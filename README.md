# MS365Apps

Installer package for MIcrosoft 365 App Suite (Word, Excel etc), Project and Visio

The ODT is pulled from a static URL, which with the appropriate XML config file runs setup for the requested components

The package behaviour can be tailored to install specific components using a parameter, ProductType, to deploy-application.exe or deploy-application.ps1
eg.

.\Deploy-Application.exe  -DeploymentType Install -DeployMode Noninteractive -ProductType 4

The configured install types are :-

|Index|Name|Purpose|
|-----|----|-------|
|0|365 Suite (user,monthly)|User account licensed 365 suite including Outlook, Monthly update channel|
|1|365 Suite (user,current)|User account licensed 365 suite including Outlook, Current update channel|
|2|365 Suite (shared,monthly)|Shared device licensed 365 suite including Outlook, user still requires a license, Monthly update channel|
|3|365 Suite (shared,current)|Shared device licensed 365 suite including Outlook, user still requires a license, Current update channel|
|4|365 Suite - no outlook (shared,monthly)|Shared device licensed 365 suite without Outlook, user still requires a license. Typically used in student labs, Monthly update channel|
|5|365 Suite - no outlook (shared,current)|Shared device licensed 365 suite without Outlook, user still requires a license. Typically used in student labs, Current update channel|
|6|Project (user,monthly)|Project, user account licensed, Monthly update channel|
|7|Project (user,current)|Project, user account licensed, Current update channel|
|8|Project (shared,monthly)|Project, shared device licensed, user still requires a license, Monthly update channel|
|9|Project (shared,current)|Project, shared device licensed, user still requires a license, Current update channel|
|10|Visio (user,monthly)|Visio, user account licensed, Monthly update channel|
|11|Visio (user,current)|Visio, user account licensed, Current update channel|
|12|Visio (shared,monthly)|Visio, shared device licensed, user still requires a license, Monthly update channel|
|13|Visio (shared,current)|Visio, shared device licensed, user still requires a license, Current update channel|



