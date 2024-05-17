# Generate-Changelog.ps1
$changelogPath = "CHANGELOG.md"

# Header for the changelog
$header = @"
# Changelog

All notable changes to this project will be documented in this file.

"@

# Write header to changelog
$header | Out-File $changelogPath

# Get the commit messages using GitVersion
$commits = & dotnet gitversion /output json | ConvertFrom-Json

# Write version and commits to changelog
$version = $commits.SemVer
"## Version $version" | Out-File -Append $changelogPath

$commitMessages = git log --pretty=format:"%h - %s (%an, %ad)" --date=short
$commitMessages | Out-File -Append $changelogPath
