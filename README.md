# Microsoft 365 Graph API - Add User
Powershell script utilizing Graph API to create M365 users without powershell modules.

The intention is for a service desk to create Microsoft 365 Accounts using base Windows Powershell 5.1.
All actions are executed using standared Invoke-RestMethod to call Graph API without requiring powershell modules to be installed.
      
An Identity Service Principal account is required in Microsoft Entra ID with the Permissiones Required:
  <ul>
    <li>Directory.ReadWrite.All</li>
    <li>UserAuthenticationMethod.ReadWrite.All</li>
</ul>
      
Actions taken by this script:
<ul>
  <li>create user account</li>
  <li>assign manager</li>
  <li>assign groups</li>
  <li>assign authentication phone number</li>
  <li>assign M365 license</li>
</ul>
By default Microsoft Entra ID Service Principals authentication tokens are valid for only one hour.
This script automatically refreshes the token when run if the Powershel ISE session remains open.

Input for this script was designed to work with tickets created from DeskDirector, but the parser could be modified to work with input from other sources.

```
Example Input:  
    ### Employee's First Name  
        John  
    ### Employee's Last Name  
        Doe  
    ### Manager Name  
        Jane Doe  
    ### Mobile Phone Number  
        +1 (555) 555-5555  
    ### Job Title  
        Technician  
    ### Department  
        Engineering  
    ### Office  
        Los Angeles  
    ### M365 Teams/Groups/Distribution Lists  
        Engineering Team, CNC Group, etc  
```

