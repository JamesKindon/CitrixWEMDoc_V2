# CitrixWEMDoc_V2

.SYNOPSIS
This script documents the Citrix Workspace Environment Management Solution.

Data is gathered using Arjan Mensch's Citrix.WEMSDK PowerShell Module

Output is sent to both Word and HTML format using the PSCribo PowerShell Module

.DESCRIPTION

.PARAMETER DBServer
Mandatory parameter specifying your SQL Database Server name or instance

.PARAMETER DBName
Mandatory parameter specifying your Citrix WEM Database Name

.PARAMETER Site
Specifies the WEM Configuration Set to document via Site ID. Defaults to Site ID 1 (Default Site)

If you do not know your site ID, use the -listAllConfigSets parameter

.PARAMETER ListAllConfigSets
Optional Parameter which can only be used with DBServer and DBName Params. Creates an initial connection to the WEM Database and lists all Configuration Sets

.PARAMETER CompanyName
Optional Parameter used to personalise the Document Output for a particular Customer Name

.PARAMETER Detailed
Optional Parameter which will output an appendix with full details for applications,rules etc

.PARAMETER OutputLocation
Optional Parameter allowing you so specify a custom output directory. Defaults to ~\Desktop

.EXAMPLE
Documents WEM based on the default SQL instance found on SERVER against the Database named WEM. Default Site (1) is used.

.\DocumentWEM_V2.ps1 -DBServer SERVER -DBName CitrixWEM
.EXAMPLE
Documents WEM based on the SQL instance named INSTANCE found on SERVER against the Database named WEM. Default Site (1) is used.

.\DocumentWEM_V2.ps1 -DBServer SERVER\INSTANCE -DBName CitrixWEM
.EXAMPLE
Lists all Config Sets in WEM based on the SQL instance named INSTANCE found on SERVER against the Database named WEM.

.\DocumentWEM_V2.ps1 -DBServer SERVER\INSTANCE -DBName CitrixWEM -ListAllConfigSets
.EXAMPLE
Documents WEM based on the SQL instance named INSTANCE found on SERVER against the Database named WEM. Site 2 is used.

.\DocumentWEM_V2.ps1 -DBServer SERVER\INSTANCE -DBName CitrixWEM -Site 2
.EXAMPLE
Documents WEM based on the SQL instance named INSTANCE found on SERVER against the Database named WEM. Default Site (1) is used. A detailed Appendix is added

.\DocumentWEM_V2.ps1 -DBServer SERVER\INSTANCE -DBName CitrixWEM -Detailed
.EXAMPLE
Documents WEM based on the SQL instance named INSTANCE found on SERVER against the Database named WEM. Default Site (1) is used. A detailed Appendix is added. The Output Location is specified

.\DocumentWEM_V2.ps1 -DBServer SERVER\INSTANCE -DBName CitrixWEM -Detailed -OutPutLocation "C:\Temp\Output"
.EXAMPLE
Documents WEM based on the SQL instance named INSTANCE found on SERVER against the Database named WEM. A detailed Appendix is added. Default Site (1) is used. a Company Name is Added

.\DocumentWEM_V2.ps1 -DBServer SERVER\INSTANCE -DBName CitrixWEM -Detailed -CompanyName "KindonEnterprises"

.NOTES
Credits as follows:
Arjan Mensch. For being the PowerShell master in relation to WEM https://github.com/msfreaks 
Aaron Parker. Functions, Parsing and basic powershell guidance throughout the project https://github.com/aaronparker/
Iain Brighton. For PSCribo https://github.com/iainbrighton/PScribo
