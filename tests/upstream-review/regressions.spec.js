describe("reviewed upstream regressions", () => {
    const w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let container;
    beforeEach(() => {
        container = document.createElement("div");
        document.body.appendChild(container);
    });
    afterEach(() => container.remove());

    const text = value => `<w:r><w:t>${value}</w:t></w:r>`;
    const para = (value, props = "") => `<w:p><w:pPr>${props}</w:pPr>${text(value)}</w:p>`;
    const escape = value => value.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
    async function render(body, { styles, settings, numbering, relationships = "", options = {} } = {}) {
        const zip = new JSZip();
        zip.file("_rels/.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${rel}/officeDocument" Target="word/document.xml"/></Relationships>`);
        const parts = { styles, settings, numbering };
        for (const [name, value] of Object.entries(parts)) {
            if (value !== undefined) {
                zip.file(`word/${name}.xml`, `<w:${name} xmlns:w="${w}">${value}</w:${name}>`);
                relationships += `<Relationship Id="${name}" Type="${rel}/${name}" Target="${name}.xml"/>`;
            }
        }
        zip.file("word/_rels/document.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`);
        zip.file("word/document.xml", `<w:document xmlns:w="${w}" xmlns:r="${rel}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440"/></w:sectPr></w:body></w:document>`);
        return docx.renderAsync(await zip.generateAsync({ type: "arraybuffer" }), container, null, options);
    }

    it("renders fragment strings, descriptors, nested fragments and existing nodes", () => {
        const node = document.createElement("i");
        node.textContent = "node";
        const fragment = docx.defaultOptions.h({ tagName: "#fragment", children: [
            "text", { tagName: "b", children: ["bold"] },
            { tagName: "#fragment", children: [node] }
        ] });
        expect(fragment.nodeType).toBe(Node.DOCUMENT_FRAGMENT_NODE);
        expect(fragment.textContent).toBe("textboldnode");
        expect(fragment.lastChild).toBe(node);
        expect(docx.defaultOptions.h({ tagName: "#fragment" }).childNodes.length).toBe(0);
    });

    it("renders visible non-breaking hyphens", async () => {
        await render(`<w:p><w:r><w:t>F</w:t><w:noBreakHyphen/><w:t>16</w:t></w:r></w:p>`);
        expect(container.querySelector("p").textContent).toBe("F\u201116");
        expect(container.querySelector("wbr")).toBeNull();
    });

    for (const enabled of [false, true]) {
        it(`exposes paragraph IDs only when enabled (${enabled})`, async () => {
            const parsed = await render(`<w:p w14:paraId="ABC123">${text("identified")}</w:p>${para("anonymous")}`, { options: { exposeParaIds: enabled } });
            expect(parsed.documentPart.body.children[0].paraId).toBe("ABC123");
            const paragraphs = container.querySelectorAll("p");
            expect(paragraphs[0].getAttribute("data-para-id")).toBe(enabled ? "ABC123" : null);
            expect(paragraphs[1].hasAttribute("data-para-id")).toBe(false);
        });
    }

    for (const target of ["javascript:alert(1)", "JaVaScRiPt:alert(1)", "java&#x09;script:alert(1)", "data:text/html,bad", "vbscript:bad", "file:///etc/passwd"]) {
        it(`blocks unsafe hyperlink ${target}`, async () => {
            await render(`<w:p><w:hyperlink r:id="link">${text("link")}</w:hyperlink></w:p>`, {
                relationships: `<Relationship Id="link" Type="${rel}/hyperlink" Target="${target}" TargetMode="External"/>`
            });
            expect(container.querySelector("a").hasAttribute("href")).toBe(false);
            expect(container.querySelector("a").textContent).toBe("link");
        });
    }
    for (const target of ["https://example.com/path?a=1&b=2", "http://example.com", "mailto:person@example.com"]) {
        it(`preserves allowed hyperlink ${target}`, async () => {
            await render(`<w:p><w:hyperlink r:id="link">${text("link")}</w:hyperlink></w:p>`, {
                relationships: `<Relationship Id="link" Type="${rel}/hyperlink" Target="${escape(target)}" TargetMode="External"/>`
            });
            expect(container.querySelector("a").getAttribute("href")).toBe(target);
        });
    }
    it("preserves internal anchors", async () => {
        await render(`<w:p><w:hyperlink w:anchor="bookmark">${text("link")}</w:hyperlink></w:p>`);
        expect(container.querySelector("a").getAttribute("href")).toBe("#bookmark");
    });

    const breakStyle = `<w:style w:type="paragraph" w:styleId="break"><w:pPr><w:pageBreakBefore/></w:pPr></w:style>`;
    it("lets direct false override pageBreakBefore from a style", async () => {
        await render(para("first") + para("second", '<w:pStyle w:val="break"/><w:pageBreakBefore w:val="0"/>'), { styles: breakStyle });
        expect(container.querySelectorAll("section.docx").length).toBe(1);
    });
    it("uses style pageBreakBefore when direct formatting is absent", async () => {
        await render(para("first") + para("second", '<w:pStyle w:val="break"/>'), { styles: breakStyle });
        expect(container.querySelectorAll("section.docx").length).toBe(2);
    });
    it("does not create an empty first page for pageBreakBefore", async () => {
        await render(para("first", '<w:pageBreakBefore/>'));
        expect(container.querySelectorAll("section.docx").length).toBe(1);
    });
    it("honors breakPages false for direct and style page breaks", async () => {
        await render(para("first") + para("second", '<w:pageBreakBefore/>') + para("third", '<w:pStyle w:val="break"/>'), { styles: breakStyle, options: { breakPages: false } });
        expect(container.querySelectorAll("section.docx").length).toBe(1);
    });

    for (const fallback of ["", "<mc:Fallback/>", '<mc:Choice Requires="w"/>']) {
        it(`handles AlternateContent with no usable child (${fallback})`, async () => {
            await render(`<w:p><w:r><mc:AlternateContent>${fallback}</mc:AlternateContent><w:t>survives</w:t></w:r></w:p>`);
            expect(container.querySelector("p").textContent).toBe("survives");
        });
    }
    for (const value of ["0", "false", "off", "1", "true", "on"]) {
        it(`parses explicit evenAndOddHeaders=${value}`, async () => {
            const parsed = await render(para("settings"), { settings: `<w:evenAndOddHeaders w:val="${value}"/>` });
            expect(parsed.settingsPart.settings.evenAndOddHeaders).toBe(["1", "true", "on"].includes(value));
        });
    }

    const table = props => `<w:tbl><w:tblPr>${props}</w:tblPr><w:tr><w:tc>${para("cell")}</w:tc></w:tr></w:tbl>`;
    for (const type of ["pct", "auto", "nil"]) {
        it(`does not interpret table indent type=${type} as twips`, async () => {
            await render(table(`<w:tblInd w:w="1440" w:type="${type}"/>`));
            expect(container.querySelector("table").style.marginInlineStart).toBe("");
        });
    }
    for (const alignment of ["center", "right"]) {
        it(`does not indent ${alignment}-aligned tables`, async () => {
            await render(table(`<w:tblInd w:w="1440" w:type="dxa"/><w:jc w:val="${alignment}"/>`));
            expect(container.querySelector("table").style.marginInlineStart).toBe("");
        });
    }
    it("retains negative table indents", async () => {
        await render(table('<w:tblInd w:w="-720" w:type="dxa"/>'));
        expect(container.querySelector("table").style.marginInlineStart).toBe("-36pt");
    });

    for (const properties of ['<w:u w:val="single"/><w:strike/>', '<w:strike/><w:u w:val="single"/>']) {
        it(`combines decorations regardless of element order (${properties})`, async () => {
            await render(`<w:p><w:r><w:rPr>${properties}</w:rPr><w:t>both</w:t></w:r></w:p>`);
            const decoration = container.querySelector("p span").style.textDecorationLine;
            expect(decoration).toContain("underline");
            expect(decoration).toContain("line-through");
        });
    }
    for (const properties of ['<w:u w:val="single"/><w:strike w:val="0"/>', '<w:strike w:val="0"/><w:u w:val="single"/>']) {
        it("turning strike off preserves the underline", async () => {
            await render(`<w:p><w:r><w:rPr>${properties}</w:rPr><w:t>underlined</w:t></w:r></w:p>`);
            expect(container.querySelector("p span").style.textDecorationLine).toBe("underline");
        });
    }
    it("turning underline off preserves strike", async () => {
        await render('<w:p><w:r><w:rPr><w:strike/><w:u w:val="none"/></w:rPr><w:t>struck</w:t></w:r></w:p>');
        expect(container.querySelector("p span").style.textDecorationLine).toBe("line-through");
    });

    it("keeps a VML textbox when imagedata has no relationship and retains wrap metadata", async () => {
        const parsed = await render('<w:p><w:r><w:pict><v:shape style="width:100pt;height:30pt"><v:imagedata/><w10:wrap type="topAndBottom"/><v:textbox><w:txbxContent>' + para("textbox") + '</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>');
        const shape = parsed.documentPart.body.children[0].children[0].children[0].children[0];
        expect(shape.wrapType).toBe("topAndBottom");
        expect(container.querySelector("svg image")).toBeNull();
        expect(container.querySelector("foreignObject").textContent).toBe("textbox");
    });

    const numberings = `<w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/></w:lvl>
        <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/></w:lvl>
        <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl>
        <w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3%4. 100%"/></w:lvl>
        </w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`;
    const numPr = level => `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="1"/></w:numPr>`;
    it("resolves referenced numbering formats without rendering bullet counters or eating literal percent", async () => {
        await render(para("nested", numPr(3)), { numbering: numberings });
        const content = getComputedStyle(container.querySelector("p"), "::before").content;
        expect(content).toContain("counter(docx-num-1-0, upper-roman)");
        expect(content).toContain("counter(docx-num-1-1)");
        expect(content).not.toContain("docx-num-1-2");
        expect(content).toContain("100%");
    });
    it("keeps numbering with text when a bookmark precedes a page break", async () => {
        await render(`<w:p><w:pPr>${numPr(0)}</w:pPr><w:bookmarkStart w:id="0" w:name="start"/><w:r><w:br w:type="page"/><w:t>after</w:t></w:r></w:p>`, { numbering: numberings });
        const paragraphs = container.querySelectorAll("p");
        expect(paragraphs[0].classList.contains("docx-numbering-suppressed")).toBe(true);
        expect(paragraphs[1].classList.contains("docx-numbering-suppressed")).toBe(false);
    });
    it("does not leave a list marker with hidden tracked deletions", async () => {
        await render(`<w:p><w:pPr>${numPr(0)}</w:pPr><w:del w:id="1"><w:r><w:delText>deleted</w:delText></w:r></w:del><w:r><w:br w:type="page"/><w:t>after</w:t></w:r></w:p>`, { numbering: numberings });
        const paragraphs = container.querySelectorAll("p");
        expect(paragraphs[0].textContent).toBe("");
        expect(paragraphs[0].classList.contains("docx-numbering-suppressed")).toBe(true);
        expect(paragraphs[1].classList.contains("docx-numbering-suppressed")).toBe(false);
    });
    it("keeps accepted inserted text and hides deleted metadata with renderChanges off", async () => {
        await render(`<w:p><w:ins w:id="0" w:author="Author">${text("accepted")}</w:ins><w:del w:id="1"><w:r><w:delText>removed</w:delText></w:r></w:del></w:p>`);
        expect(container.querySelector("p").textContent).toBe("accepted");
        expect(container.querySelector("[data-change-id], ins, del")).toBeNull();
    });
    it("does not invent missing change metadata", async () => {
        await render(`<w:p><w:ins>${text("accepted")}</w:ins></w:p>`, { options: { renderChanges: true } });
        const ins = container.querySelector("ins");
        expect(ins.textContent).toBe("accepted");
        expect(ins.hasAttribute("data-change-author")).toBe(false);
        expect(ins.hasAttribute("data-change-date")).toBe(false);
        expect(ins.hasAttribute("data-change-id")).toBe(false);
    });

    it("keeps text before a break at the end of a run on its original page", async () => {
        await render(`<w:p><w:pPr>${numPr(0)}</w:pPr><w:r><w:t>before</w:t><w:br w:type="page"/></w:r>${text("after")}</w:p>`, { numbering: numberings });
        const paragraphs = container.querySelectorAll("p");
        expect(paragraphs[0].textContent).toBe("before");
        expect(paragraphs[1].textContent).toBe("after");
        expect(paragraphs[0].classList.contains("docx-numbering-suppressed")).toBe(false);
        expect(paragraphs[1].classList.contains("docx-numbering-suppressed")).toBe(true);
    });

    it("uses inherited tab stops before rendering and measures successive tabs against updated positions", async () => {
        await render('<w:p><w:pPr><w:pStyle w:val="tabs"/></w:pPr>' + text("A") + '<w:r><w:tab/></w:r>' + text("B") + '<w:r><w:tab/></w:r>' + text("C") + '</w:p>', {
            styles: '<w:style w:type="paragraph" w:styleId="tabs"><w:pPr><w:tabs><w:tab w:val="left" w:pos="2880"/><w:tab w:val="left" w:pos="5760"/></w:tabs></w:pPr></w:style>',
            settings: '<w:defaultTabStop w:val="720"/>', options: { experimental: true }
        });
        await new Promise(resolve => setTimeout(resolve, 650));
        const p = container.querySelector("p");
        const spans = [...p.querySelectorAll("span")];
        const offset = value => spans.find(s => s.textContent === value).getBoundingClientRect().left - p.getBoundingClientRect().left;
        expect(Math.abs(offset("B") - 192)).toBeLessThan(8);
        expect(Math.abs(offset("C") - 384)).toBeLessThan(8);
    });
});
