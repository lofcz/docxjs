#!/usr/bin/env bash
set -euo pipefail

mode=publish
if (( $# > 1 )); then
    echo 'Usage: ./bootstrap-publish.sh [--dry-run | --trust-only]' >&2
    exit 2
fi
case "${1:-}" in
    '') ;;
    --dry-run) mode=dry-run ;;
    --trust-only) mode=trust-only ;;
    --help|-h)
        echo 'Usage: ./bootstrap-publish.sh [--dry-run | --trust-only]'
        exit 0 ;;
    *) echo "Unknown option: $1" >&2; exit 2 ;;
esac

cd -- "$(dirname -- "${BASH_SOURCE[0]}")"
registry=https://registry.npmjs.org/
node <<'JS'
const pkg = require('./package.json');
if (pkg.name !== '@lofcz/docx-preview' || pkg.repository.url !== 'git+https://github.com/lofcz/docxjs.git') {
    throw new Error('Package metadata must identify @lofcz/docx-preview and lofcz/docxjs.');
}
const [major, minor] = process.versions.node.split('.').map(Number);
if (major < 22 || (major === 22 && minor < 14)) {
    throw new Error('Install Node.js 22.14+ (Node.js 24 LTS recommended).');
}
JS
npm_version=$(npm --version)
node - "$npm_version" <<'JS'
const [major, minor] = process.argv[2].split('.').map(Number);
if (!Number.isInteger(major) || major < 11 || (major === 11 && minor < 15)) {
    throw new Error('npm trust requires npm 11.15+. Run: npm install --global npm@11.19.0');
}
JS
package_name=$(node -p "require('./package.json').name")
package_version=$(node -p "require('./package.json').version")

if [[ "$mode" != trust-only ]]; then
    # Karma normally discovers Chrome; also support common Linux Chromium names.
    if [[ -z "${CHROME_BIN:-}" ]]; then
        for browser in google-chrome google-chrome-stable chromium chromium-browser; do
            if browser_path=$(command -v "$browser"); then
                export CHROME_BIN="$browser_path"
                break
            fi
        done
    fi
    echo 'Installing dependencies and validating the release...'
    npm ci
    npm run verify-release
    npm pack --dry-run
fi

if [[ "$mode" == dry-run ]]; then
    npm publish --dry-run --access public --registry="$registry"
    echo 'Dry run complete. No login, publish or trust changes were performed.'
    exit 0
fi

echo 'Log in to npm with an account that can publish to @lofcz (2FA required).'
npm login --auth-type=web --registry="$registry"
if [[ "$mode" != trust-only ]]; then
    npm publish --access public --registry="$registry"
fi

spec="$package_name@$package_version"
visible=false
for (( attempt=1; attempt<=12; attempt++ )); do
    if published_version=$(npm view "$spec" version --registry="$registry" 2>/dev/null) &&
        [[ "$published_version" == "$package_version" ]]; then
        visible=true
        break
    fi
    if (( attempt < 12 )); then sleep 5; fi
done
if [[ "$visible" != true ]]; then
    echo "$spec is not visible yet. Once visible, rerun with --trust-only." >&2
    exit 1
fi

echo 'Authorizing release.yml as the npm trusted publisher...'
npm trust github "$package_name" --repo lofcz/docxjs --file release.yml \
    --allow-publish --yes --registry="$registry"
npm trust list "$package_name" --registry="$registry"
echo 'Bootstrap complete. Future releases use release.yml with OIDC; no NPM_TOKEN is needed.'
