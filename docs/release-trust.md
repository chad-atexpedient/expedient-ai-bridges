# Release trust and provenance

## Unsigned reproducibility boundary

CI packages only allowlisted committed files and Node distributions pinned by SHA-256 in `eng/release-inputs.json`. The Git commit timestamp is the source epoch; local filesystem timestamps are ignored. The release manifest records commit, RID, logical paths, sizes, and hashes. Compiled bridge bundles are exact attested inputs, not claimed source-reproducible builds.

## Production trust boundary

A production release requires a protected `production-release` GitHub Environment and fails closed unless all organization-controlled credentials are configured:

- Windows organization signing through Azure Trusted Signing or equivalent, SHA-256 plus RFC 3161 timestamp;
- Apple Developer ID Application identity;
- App Store Connect/notarytool credential;
- required human reviewer approval.

Nested macOS code is signed inside-out: Node, companion, application bundle, then DMG. CI verifies `codesign --deep --strict`, Gatekeeper, notarization, and stapling. Windows CI verifies Authenticode status and timestamp. Final hashes are generated only after these transformations.

## Evidence shipped together

Every candidate contains a release manifest, sorted SHA-256 file, and SPDX JSON SBOM. Final signed subjects—not pre-sign digests—receive GitHub artifact attestations using OIDC. Signing timestamps and notarization tickets are intentionally nondeterministic and are mapped to the reproducible unsigned input manifest.

## Current credential status

The repository contains the controls and fail-closed workflow, not private organization credentials. Until the repository owner configures the protected environment and obtains certificates, CI may publish explicitly named unsigned candidates for engineering validation, but cannot publish a normal signed production release.
