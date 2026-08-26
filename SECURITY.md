# Security policy

## Reporting

Do not open a public issue containing credentials, private endpoints, customer data, certificate material, or exploit details. Report security concerns privately through the repository's GitHub Security Advisory interface.

## Credential boundary

Only templates and placeholders belong in Git. Never commit:

- `config.env` or provider/API credentials;
- OAuth/Graph tokens or client secrets;
- signing certificates, private keys, notarization keys, or keychain exports;
- generated localhost TLS keys/certificates;
- logs, telemetry, PID/control state, or user-generated Office files.

The repository runs Gitleaks against the complete reachable Git history on pushes, pull requests, schedules, and manual dispatches. Release workflows repeat the scan before packaging.

## If a credential is exposed

1. Revoke or rotate it at the issuing provider immediately. History rewriting is not revocation.
2. Remove the file from the current tree and add an ignore rule.
3. Use `git filter-repo` or an equivalent reviewed procedure to remove the value from every affected ref when required.
4. Force-update only after coordinating with every clone/fork owner.
5. Re-run full-history secret scanning and document the incident privately.
6. Rotate any related credentials that shared scope or trust.

No credential discovered locally during packaging should ever be copied into an issue, build log, report, or chat transcript.
