# Publishing the fork

This fork publishes `@lofcz/docx-preview` from `lofcz/docxjs`. The initial version is `0.4.0`; the package name is separate from upstream's `docx-preview`.

## First publication

Push the release setup to `master` before publishing. Use Node.js 24, PowerShell 7, and an npm account with access to the `@lofcz` scope and two-factor authentication enabled.

```powershell
npm install --global npm@11.19.0
pwsh ./bootstrap-publish.ps1 -DryRun
pwsh ./bootstrap-publish.ps1
```

The script installs locked dependencies, typechecks, builds all distribution bundles, runs the headless Chrome tests, and lists the package contents. Install Chrome locally, or set `CHROME_BIN` to a Chromium executable before running it.

The actual bootstrap then runs these commands, with interactive npm login and 2FA:

```sh
npm login --auth-type=web --registry=https://registry.npmjs.org/
npm publish --access public --registry=https://registry.npmjs.org/
npm trust github @lofcz/docx-preview --repo lofcz/docxjs --file release.yml --allow-publish --yes --registry=https://registry.npmjs.org/
npm trust list @lofcz/docx-preview --registry=https://registry.npmjs.org/
```

The script waits for the exact version to appear before configuring trust. If publication succeeds but trust setup fails, rerun `pwsh ./bootstrap-publish.ps1 -TrustOnly`. This logs in and configures trust without republishing. An existing trust configuration is never automatically revoked; inspect it with `npm trust list` if npm reports that one already exists.

`-DryRun` performs validation and an npm publish dry run, without login or registry writes. The script does not store tokens in the repository. The initial local publication has no GitHub provenance; subsequent OIDC releases do.

[npm trust](https://docs.npmjs.com/cli/v11/commands/npm-trust/) requires npm 11.15+, an existing package and account-level 2FA. The trusted publisher is bound to `lofcz/docxjs` and the filename `release.yml`, with no GitHub environment configured.

## Subsequent releases

Run **Release and Publish to npm** in GitHub Actions on `master`, choosing `patch`, `minor` or `major`. For example:

```sh
gh workflow run release.yml --repo lofcz/docxjs --ref master -f release_type=patch
```

The workflow uses Node.js 24 and npm 11.19.0, installs with `npm ci`, bumps both manifests, builds and tests, then commits the manifests and generated `dist` files. It pushes that commit and its annotated tag atomically before publishing with OIDC and creating a GitHub release. Only one release runs at a time. No npm token secret is required; `id-token: write` enables [npm trusted publishing](https://docs.npmjs.com/trusted-publishers/).

If npm publication fails after the commit/tag push, fix the cause and dispatch `release_type=current` while `master` still points at that release commit. This retries the same version. Never move a published version's tag. If npm publication succeeded but GitHub release creation failed, create only the missing GitHub release:

```sh
gh release create v0.4.1 --repo lofcz/docxjs --verify-tag --generate-notes --title v0.4.1
```

Use the actual published version in place of `0.4.1`. npm will reject publishing an existing version; the workflow does not treat that rejection as success. The workflow needs permission to push to `master`; repository branch protection must permit its release commits.
