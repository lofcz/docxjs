[CmdletBinding()]
param(
    [switch]$DryRun,
    [switch]$TrustOnly
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Invoke-Npm {
    param([string[]]$Arguments)
    & npm @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "npm $($Arguments -join ' ') failed with exit code $LASTEXITCODE"
    }
}

if ($DryRun -and $TrustOnly) {
    throw 'Use either -DryRun or -TrustOnly, not both.'
}

Push-Location $PSScriptRoot
try {
    $package = Get-Content -Raw (Join-Path $PSScriptRoot 'package.json') | ConvertFrom-Json
    $registry = 'https://registry.npmjs.org/'
    if ($package.name -ne '@lofcz/docx-preview' -or
        $package.repository.url -ne 'git+https://github.com/lofcz/docxjs.git') {
        throw 'Package metadata must identify @lofcz/docx-preview and lofcz/docxjs.'
    }

    $nodeVersion = & node -p 'process.versions.node'
    if ($LASTEXITCODE -ne 0 -or [version]$nodeVersion -lt [version]'22.14.0') {
        throw 'Install Node.js 22.14+ (Node.js 24 LTS recommended).'
    }
    $npmVersion = & npm --version
    if ($LASTEXITCODE -ne 0 -or [version]$npmVersion -lt [version]'11.15.0') {
        throw 'npm trust requires npm 11.15+. Run: npm install --global npm@11.19.0'
    }

    if (-not $TrustOnly) {
        Write-Host 'Installing dependencies and validating the release...'
        Invoke-Npm -Arguments @('ci')
        Invoke-Npm -Arguments @('run', 'verify-release')
        Invoke-Npm -Arguments @('pack', '--dry-run')
    }

    if ($DryRun) {
        Invoke-Npm -Arguments @('publish', '--dry-run', '--access', 'public', "--registry=$registry")
        Write-Host 'Dry run complete. No login, publish or trust changes were performed.'
    }
    else {
        Write-Host 'Log in to npm with an account that can publish to @lofcz (2FA required).'
        Invoke-Npm -Arguments @('login', '--auth-type=web', "--registry=$registry")

        if (-not $TrustOnly) {
            # Initial local publication cannot generate GitHub Actions provenance.
            Invoke-Npm -Arguments @('publish', '--access', 'public', "--registry=$registry")
        }

        $spec = "$($package.name)@$($package.version)"
        $visible = $false
        for ($attempt = 1; $attempt -le 12; $attempt++) {
            $publishedVersion = & npm view $spec version "--registry=$registry" 2>$null
            if ($LASTEXITCODE -eq 0 -and "$publishedVersion".Trim() -eq $package.version) {
                $visible = $true
                break
            }
            if ($attempt -lt 12) { Start-Sleep -Seconds 5 }
        }
        if (-not $visible) {
            throw "$spec is not visible yet. Once visible, rerun with -TrustOnly."
        }

        Write-Host 'Authorizing release.yml as the npm trusted publisher...'
        Invoke-Npm -Arguments @('trust', 'github', $package.name,
            '--repo', 'lofcz/docxjs', '--file', 'release.yml', '--allow-publish', '--yes',
            "--registry=$registry")
        Invoke-Npm -Arguments @('trust', 'list', $package.name, "--registry=$registry")
        Write-Host 'Bootstrap complete. Future releases use release.yml with OIDC; no NPM_TOKEN is needed.'
    }
}
finally {
    Pop-Location
}
