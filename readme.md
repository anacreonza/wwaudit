# Woodwing Report Powershell

This script creates a report in JSON format of the state and configuration of the publishing software installed on this PC.

It checks:

1. Computer name and user name.
2. The last modified date of the WWSettings.xml and ApplicationMapping.xml files that are required by Woodwing software.
3. The version of Content Station AIR that is installed on the PC.
4. It does a search of the registry for installed Adobe products.
5. It checks the version of the Smart Connection plugins installed on the PC.
6. It generates a list of the drive mappings of the PC.