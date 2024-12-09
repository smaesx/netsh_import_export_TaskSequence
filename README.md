# netsh_import_export_TaskSequence
Exporting and importing of ethernet network adapter settings within MECM Task Sequence Environment.

## Backup_netsh_adapter.ps1
Exports the ethernet adapter configuration to `$logDirectory\<NICName>.xml`.

Transcript stored in `$logDirectory\Netsh_Export.log`.

**Note**: You need to adjust the `$logDirectory` variable at line 8.

## Import_netsh_adapter.ps1
Imports the ethernet adapter configuration from `$logDirectory\<NICName>.xml` and disables/enables the network adapter.

Transcript stored in `$logDirectory\Netsh_Import.log`.

**Note**: You need to adjust the `$logDirectory` variable at line 8.
