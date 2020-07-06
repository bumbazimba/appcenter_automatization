### AppCenter automated builds with report ###

**Technologies:** | Powershell | AppCenter | API 
---|---|---|---

This Powershell script allows user to build all branches for a selected Application in [App Center] and get report in Excel format. Script based on App Center [API] requests.
### Pre-requirments: ###

1 Registered account in AppCenter  
2 Existing Application with configured branches. If branch is not configured, user may get error  
3 User API token  

### Execution details ###

1 User will be prompted to Enter following data:  
  - AppCenter account name  
  - Application name  
  - API token  
  - Number of branches to build  
  
2 Excel file template will be created  
3 All branches for an Application will be listed  
4 First several branches will be queued for building (number depends on user's input from step 1)  
5 When a build completed, for each build gathering information about build result (failed or success), Duration and Log file link  
6 Saving all details in Excel file on User's Desktop  


### Notes: ###  
If you are using a free plan for an AppCenter, check out [limitations]


[App Center]: https://appcenter.ms/
[API]: https://openapi.appcenter.ms/#/
[limitations]: https://docs.microsoft.com/en-us/appcenter/general/pricing
