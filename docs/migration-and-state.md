# Installation, migration, rollback, and state ownership

## One owner per profile

The source-built `ExpedientBridges` supervisor is the sole owner of bridge child processes for a migrated profile. It owns loopback listeners 4001 and 4002, the supervisor lock/state/stop request, bounded child restart, and redacted rotating logs. Legacy PowerShell and packaged v1.1 companions must not run concurrently with the v1.2 supervisor.

## Mutable state contract

- Windows: `%LOCALAPPDATA%\ExpedientAIBridges`
- macOS: `~/Library/Application Support/Expedient AI Bridges`

`config.env` is preserved across install, upgrade, migration, and uninstall unless the user deliberately purges data. Executable payloads are immutable and versioned separately.

## Upgrade to schema 2

1. Stop the legacy script edition and packaged companion.
2. Verify ports 4001 and 4002 have no unrelated listeners.
3. Install the v1.2 payload and run `ExpedientBridges migrate`.
4. Migration preserves configuration, renames legacy PID records to `*.legacy`, and writes `runtime/migration-v2.json`.
5. Run `ExpedientBridges start`, then `status` and `diagnostics`.

Startup also performs this idempotent migration. Unknown port owners are never killed.

## Upgrade and rollback

Stop the supervisor before replacing immutable files. Stage a complete new version before activation and preserve mutable state. To roll back, stop and verify closed ports, restore the prior immutable payload, keep configuration, and never restore stale PID files. Legacy PID backups are audit evidence only.

## Cleanup

Uninstall removes immutable program files and shortcuts. Mutable state remains by default for reinstall or rollback. Purge it only after shutdown and any desired configuration backup.
