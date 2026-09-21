describe("fork network regressions", () => {
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
    async function render(body, { styles, settings, numbering, relationships = "", files = {}, options = {} } = {}) {
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
        for (const [path, content] of Object.entries(files)) zip.file(path, content);
        return docx.renderAsync(await zip.generateAsync({ type: "arraybuffer" }), container, null, options);
    }

    for (const [value, expected] of [["", true], [' w:val="false"', false], [' w:val="1"', true]]) {
        it(`parses autoHyphenation on/off (${value})`, async () => {
            const parsed = await render(para("hyphens"), { settings: `<w:autoHyphenation${value}/>` });
            expect(parsed.settingsPart.settings.autoHyphenation).toBe(expected);
        });
    }
    it("keeps absent autoHyphenation unspecified", async () => {
        const parsed = await render(para("text"), { settings: "" });
        expect(parsed.settingsPart.settings.autoHyphenation).toBeUndefined();
    });
    it("retains comment boundaries inside hyperlinks", async () => {
        const parsed = await render(`<w:p><w:hyperlink w:anchor="target"><w:commentRangeStart w:id="1"/>${text("link")}<w:commentRangeEnd w:id="1"/></w:hyperlink></w:p>`);
        const children = parsed.documentPart.body.children[0].children[0].children;
        expect(children.map(x => x.type)).toEqual(["commentRangeStart", "run", "commentRangeEnd"]);
        expect(container.querySelector("a").textContent).toBe("link");
    });
    for (const reverse of [false, true]) {
        for (const color of ["yellow", "none"]) {
            it(`resolves highlight ${color} over shading in either order (${reverse})`, async () => {
                let props = ['<w:highlight w:val="' + color + '"/>', '<w:shd w:fill="FF0000"/>'];
                if (reverse) props.reverse();
                await render(`<w:p><w:r><w:rPr>${props.join("")}</w:rPr><w:t>colored</w:t></w:r></w:p>`);
                expect(getComputedStyle(container.querySelector("p span")).backgroundColor).toBe(color === "yellow" ? "rgb(255, 255, 0)" : "rgb(255, 0, 0)");
            });
        }
    }
    it("keeps document defaults but ignores an ordinary style without an ID", async () => {
        await render(para("default"), { styles: `<w:docDefaults><w:rPrDefault><w:rPr><w:color w:val="0000FF"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="character"><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>` });
        expect(getComputedStyle(container.querySelector("p span")).color).toBe("rgb(0, 0, 255)");
    });
    for (const color of ["#4472c4 [3204]", "red", "none"]) {
        it(`normalizes VML colors without losing ${color}`, async () => {
            const parsed = await render(`<w:p><w:r><w:pict><v:rect fillcolor="${color}"><v:stroke color="${color}"/></v:rect></w:pict></w:r></w:p>`);
            const shape = parsed.documentPart.body.children[0].children[0].children[0].children[0];
            expect(shape.attrs.fill).toBe(color.split(" ")[0]);
            expect(shape.attrs.stroke).toBe(color.split(" ")[0]);
        });
    }
    it("renders cached simple fields, including nested fields and formatting", async () => {
        const parsed = await render(`<w:p><w:fldSimple w:instr="PAGE" w:fldLock="true"><w:r><w:rPr><w:b/></w:rPr><w:t>7</w:t></w:r><w:fldSimple w:instr="NUMPAGES">${text("12")}</w:fldSimple></w:fldSimple></w:p>`);
        expect(container.querySelector("p").textContent).toBe("712");
        expect(getComputedStyle(container.querySelector("p span")).fontWeight).toBe("700");
        expect(parsed.documentPart.body.children[0].children[0].lock).toBe(true);
    });
    it("renders superscript footnote references only once", async () => {
        await render(`<w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>`, {
            relationships: `<Relationship Id="notes" Type="${rel}/footnotes" Target="footnotes.xml"/>`,
            files: { "word/footnotes.xml": `<w:footnotes xmlns:w="${w}"><w:footnote w:id="1">${para("note")}</w:footnote></w:footnotes>` }
        });
        expect(container.querySelector("article p").textContent).toBe("1");
        expect(container.querySelectorAll("section>ol>li").length).toBe(1);
    });
    it("sandboxes altChunk HTML and preserves opt-out", async () => {
        const opts = {
            relationships: `<Relationship Id="html" Type="${rel}/aFChunk" Target="chunk.html"/>`,
            files: { "word/chunk.html": '<p>cached HTML</p><script>parent.__docxChunkExecuted = true;</script>' }
        };
        window.__docxChunkExecuted = false;
        await render('<w:altChunk r:id="html"/>', opts);
        const frame = container.querySelector("iframe");
        expect(frame.hasAttribute("sandbox")).toBe(true);
        expect(frame.getAttribute("sandbox")).toBe("");
        expect(frame.srcdoc).toContain("cached HTML");
        await new Promise(resolve => setTimeout(resolve, 50));
        expect(window.__docxChunkExecuted).toBe(false);
        await render('<w:altChunk r:id="html"/>', { ...opts, options: { renderAltChunks: false } });
        expect(container.querySelector("iframe")).toBeNull();
        delete window.__docxChunkExecuted;
    });
    const numbering = `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`;
    const numPr = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';
    it("keeps quoted list labels inside a CSS string", async () => {
        const label = '"} .docx p {color: rgb(1, 2, 3)} /* %1 \\ tail';
        await render(para("safe", numPr), { numbering: numbering.replace('%1.', escape(label)) });
        expect(getComputedStyle(container.querySelector("p")).color).toBe("rgb(0, 0, 0)");
        const content = getComputedStyle(container.querySelector("p"), "::before").content;
        expect(content).toContain('counter(docx-num-1-0');
        expect(content).toContain('tail');
    });
    it("uses the declared default paragraph style for numbering and direct false precedence", async () => {
        await render(para("one") + para("two", '<w:pageBreakBefore w:val="false"/>'), {
            numbering,
            styles: `<w:style w:type="paragraph" w:styleId="Base" w:default="1"><w:pPr>${numPr}<w:pageBreakBefore/></w:pPr></w:style>`
        });
        expect(container.querySelectorAll("section.docx").length).toBe(1);
        expect(container.querySelectorAll("p.docx-num-1-0").length).toBe(2);
    });
    for (const separateRuns of [false, true]) {
        it(`splits every page break and preserves parsed content (${separateRuns})`, async () => {
            const br = '<w:br w:type="page"/>';
            const body = separateRuns ? `${text("A")}<w:r>${br}</w:r>${text("B")}<w:r>${br}</w:r>${text("C")}` : `<w:r><w:t>A</w:t>${br}<w:t>B</w:t>${br}<w:t>C</w:t></w:r>`;
            const parsed = await render(`<w:p><w:pPr>${numPr}</w:pPr>${body}</w:p>`, { numbering });
            expect([...container.querySelectorAll("article")].map(x => x.textContent)).toEqual(["A", "B", "C"]);
            expect(container.querySelectorAll("p.docx-numbering-suppressed").length).toBe(2);
            const original = parsed.documentPart.body.children[0];
            const countBreaks = e => (e.type === "break" ? 1 : 0) + (e.children ?? []).reduce((n, x) => n + countBreaks(x), 0);
            expect(countBreaks(original)).toBe(2);
            const nodes = await docx.renderDocument(parsed);
            container.replaceChildren(...nodes);
            expect([...container.querySelectorAll("article")].map(x => x.textContent)).toEqual(["A", "B", "C"]);
            expect(container.querySelectorAll("p.docx-numbering-suppressed").length).toBe(2);
        });
    }
    it("preserves consecutive empty pages and gives numbering to the first text", async () => {
        await render(`<w:p><w:pPr>${numPr}</w:pPr><w:r><w:br w:type="page"/><w:br w:type="page"/><w:t>third</w:t></w:r></w:p>`, { numbering });
        expect([...container.querySelectorAll("article")].map(x => x.textContent)).toEqual(["", "", "third"]);
        expect(container.querySelectorAll("p.docx-numbering-suppressed").length).toBe(2);
        expect(container.querySelectorAll("p")[2].classList.contains("docx-numbering-suppressed")).toBe(false);
    });
    it("keeps all content on one page when breakPages is disabled", async () => {
        await render('<w:p><w:r><w:t>A</w:t><w:br w:type="page"/><w:t>B</w:t><w:br w:type="page"/><w:t>C</w:t></w:r></w:p>', { options: { breakPages: false } });
        expect(container.querySelectorAll("section.docx").length).toBe(1);
        expect(container.querySelector("article").textContent).toBe("ABC");
    });
    it("places inherited tabs before renderAsync resolves", async () => {
        await render('<w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:tab/><w:t>C</w:t></w:r></w:p>', {
            styles: '<w:style w:type="paragraph" w:styleId="Base" w:default="1"><w:pPr><w:tabs><w:tab w:val="left" w:pos="1440"/><w:tab w:val="left" w:pos="2880"/></w:tabs></w:pPr></w:style>',
            options: { experimental: true }
        });
        const p = container.querySelector("p");
        const tabs = [...p.querySelectorAll('.docx-tab-stop')];
        expect(tabs.length).toBe(2);
        expect(Math.abs(tabs[0].getBoundingClientRect().right - p.getBoundingClientRect().left - 96)).toBeLessThan(6);
        expect(Math.abs(tabs[1].getBoundingClientRect().right - p.getBoundingClientRect().left - 192)).toBeLessThan(6);
    });
    const tableStyles = `<w:style w:type="table" w:styleId="Bands"><w:tblPr><w:tblStyleRowBandSize w:val="2"/></w:tblPr><w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:fill="FF0000"/></w:tcPr></w:tblStylePr><w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:fill="00FF00"/></w:tcPr></w:tblStylePr><w:tblStylePr w:type="band2Horz"><w:tcPr><w:shd w:fill="0000FF"/></w:tcPr></w:tblStylePr></w:style>`;
    const cell = (label, props = "") => `<w:tc><w:tcPr>${props}</w:tcPr>${para(label)}</w:tc>`;
    const row = (label, props = "") => `<w:tr><w:trPr>${props}</w:trPr>${cell(label)}</w:tr>`;
    it("derives table header and bands with inherited band sizes", async () => {
        await render(`<w:tbl><w:tblPr><w:tblStyle w:val="Bands"/><w:tblLook w:firstRow="1" w:noVBand="1"/></w:tblPr>${["head", "one", "two", "three", "four"].map(x => row(x)).join("")}</w:tbl>`, { styles: tableStyles });
        expect([...container.querySelectorAll("td")].map(x => getComputedStyle(x).backgroundColor)).toEqual([
            "rgb(255, 0, 0)", "rgb(0, 255, 0)", "rgb(0, 255, 0)", "rgb(0, 0, 255)", "rgb(0, 0, 255)"
        ]);
    });
    it("honors disabled table bands and explicit all-false conditional formatting", async () => {
        await render(`<w:tbl><w:tblPr><w:tblStyle w:val="Bands"/><w:tblLook w:noHBand="1" w:noVBand="1"/></w:tblPr>${row("none", '<w:cnfStyle w:oddHBand="1"/>')}</w:tbl>`, { styles: tableStyles });
        expect(getComputedStyle(container.querySelector("td")).backgroundColor).toBe("rgba(0, 0, 0, 0)");
        await render(`<w:tbl><w:tblPr><w:tblStyle w:val="Bands"/><w:tblLook w:firstRow="1"/></w:tblPr>${row("none", '<w:cnfStyle w:val="000000000000"/>')}</w:tbl>`, { styles: tableStyles });
        expect(getComputedStyle(container.querySelector("td")).backgroundColor).toBe("rgba(0, 0, 0, 0)");
    });
    it("does not paint nested cells with an outer table's conditional fill", async () => {
        await render(`<w:tbl><w:tblPr><w:tblStyle w:val="Bands"/><w:tblLook w:firstRow="1"/></w:tblPr><w:tr><w:tc>${para("outer")}<w:tbl>${row("inner")}</w:tbl></w:tc></w:tr></w:tbl>`, { styles: tableStyles });
        const cells = container.querySelectorAll("td");
        expect(getComputedStyle(cells[0]).backgroundColor).toBe("rgb(255, 0, 0)");
        expect(getComputedStyle(cells[1]).backgroundColor).toBe("rgba(0, 0, 0, 0)");
    });
    it("accounts for skipped grid columns and spans in conditional column classes", async () => {
        const parsed = await render(`<w:tbl><w:tblPr><w:tblLook w:firstColumn="1" w:lastColumn="1"/></w:tblPr><w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid><w:tr><w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>${cell("middle", '<w:gridSpan w:val="2"/>')}</w:tr></w:tbl>`);
        const cls = parsed.documentPart.body.children[0].children[0].children[0].className;
        expect(cls).not.toContain("first-col");
        expect(cls).not.toContain("last-col");
        expect(cls).toContain("odd-col");
    });
    it("resolves styles whose IDs coincide with object prototype keys", async () => {
        await render(para("safe", '<w:pStyle w:val="__proto__"/>'), { styles: '<w:style w:type="paragraph" w:styleId="__proto__"><w:rPr><w:color w:val="0000FF"/></w:rPr></w:style>' });
        expect(getComputedStyle(container.querySelector("p span")).color).toBe("rgb(0, 0, 255)");
    });
    it("lets explicit table look flags override the packed bitmask", async () => {
        await render(`<w:tbl><w:tblPr><w:tblStyle w:val="Bands"/><w:tblLook w:val="0020" w:firstRow="0" w:noHBand="1" w:noVBand="1"/></w:tblPr>${row("plain")}</w:tbl>`, { styles: tableStyles });
        expect(getComputedStyle(container.querySelector("td")).backgroundColor).toBe("rgba(0, 0, 0, 0)");
    });
    it("keeps direct cell shading over derived table bands", async () => {
        await render(`<w:tbl><w:tblPr><w:tblStyle w:val="Bands"/><w:tblLook w:firstRow="1"/></w:tblPr><w:tr>${cell("direct", '<w:shd w:fill="123456"/>')}</w:tr></w:tbl>`, { styles: tableStyles });
        expect(getComputedStyle(container.querySelector("td")).backgroundColor).toBe("rgb(18, 52, 86)");
    });
});
