# Apply-ZimbraPermissions

This script uses exported Zimbra permissions, and maps them to Exchange permissions to
assist in post-migration configuration of mailbox-related permissions.
	
## Prerequisites

Script requires an active connection to the target Exchange environment, either through the 
Exchange Management Shell or Exchange Online Management module.
	
## Usage

```
``Apply-Permissions.p1 -UsersFile Users.csv -PermissionsFile Permissions.csv
Applies permissions from Permissions.csv using mailboxes specified in Users.csv.

``Apply-Permissions.p1 -UsersFile Users.csv -PermissionsFolder C:\PermissionFiles -Confirm:$false 
Applies permissions using permission files located in C:\PermissionFiles, using mailboxes specified in Users.csv. Script will not ask
for confirmation for each operation.

``Apply-Permissions.p1 -UsersFile Users.csv -PermissionsFile Permissions.csv -WhatIf:$true | Export-Csv PermissionReport.csv -NoTypeInformation    
Export what permissions would be applied from Permissions.csv using mailboxes specified in Users.csv. Output is exported to PermissionReport.csv.
```

## Contributing

N/A

## Versioning

Initial version published on GitHub is 1.0. Changelog is contained in the script.

## Authors

* Michel de Rooij: https://github.com/michelderooij

## License

This project is licensed under the MIT License - see the LICENSE.md for details.

## Acknowledgments

N/A
 