using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Web.Http;
using static WebAPI.Models.MailBoxFolderPermissionModels;

namespace WebAPI.Helpers
{
    public class MailBoxFolderPermissionHelpers
    {
       
        [NonAction]

        public PowerShell PSmailBoxFolderPermission(string IdentityName, ref Runspace remoteRunspace)
        {
            try
            {
                Command activityCommand = new Command("Get-MailboxFolderPermission");
                activityCommand.Parameters.Add("-Identity", IdentityName);
                PowerShell powershell = PowerShell.Create();
                powershell.Runspace = remoteRunspace;
                powershell.Commands.AddCommand(activityCommand);
                //var results = powershell.Invoke();
                // return results;
                return powershell;
            }
            catch (System.Exception E)
            {
                remoteRunspace.Close();
                remoteRunspace.Dispose();
                throw E;
            }
        }

       
        [NonAction]

        public IEnumerable<GetMailBoxFolderPermission> RootFolderPermissionResults(ICollection<PSObject> aCollection)
        {
            return aCollection.Select(i => new GetMailBoxFolderPermission
            {
                DisplayName = ((PSObject)i.Properties["User"].Value).ToString(),
                UserPrincipalName = ((PSObject)((PSObject)i.Properties["User"].Value).Properties["ADRecipient"].Value) == null ? ((PSObject)i.Properties["User"].Value).ToString() : ((PSObject)((PSObject)i.Properties["User"].Value).Properties["ADRecipient"].Value).Properties["UserPrincipalName"].Value.ToString(),
                AliasName = ((PSObject)((PSObject)i.Properties["User"].Value).Properties["ADRecipient"].Value) == null ? ((PSObject)i.Properties["User"].Value).ToString() : ((PSObject)((PSObject)i.Properties["User"].Value).Properties["ADRecipient"].Value).Properties["Alias"].Value.ToString(),
                AccessRights = ((PSObject)(i.Properties["AccessRights"].Value)).ImmediateBaseObject
            });
        }

      
        public PowerShell AddMailBoxFolderPermission(string IdentityName,string UserName,string[] AccessRightsName, ref Runspace remoteRunspace)
        {
            try
            {
                Command activityCommand = new Command("Add-MailboxFolderPermission");
                activityCommand.Parameters.Add("-Identity", IdentityName);
                activityCommand.Parameters.Add("-User", UserName);
                activityCommand.Parameters.Add("-AccessRights", AccessRightsName);

                PowerShell powershell = PowerShell.Create();
                powershell.Runspace = remoteRunspace;
                powershell.Commands.AddCommand(activityCommand);
                return powershell;
            }
            catch (System.Exception E)
            {
                remoteRunspace.Close();
                remoteRunspace.Dispose();
                throw E;
            }
        }

       
        [NonAction]

        public PowerShell SetMailBoxFolderPermission(string IdentityName, string UserName, string[] AccessRightsName, ref Runspace remoteRunspace)
        {
            try
            {
                Command activityCommand = new Command("Set-MailboxFolderPermission");
                activityCommand.Parameters.Add("-Identity", IdentityName);
                activityCommand.Parameters.Add("-User", UserName);
                activityCommand.Parameters.Add("-AccessRights", AccessRightsName);

                PowerShell powershell = PowerShell.Create();
                powershell.Runspace = remoteRunspace;
                powershell.Commands.AddCommand(activityCommand);
                return powershell;
            }
            catch (System.Exception E)
            {
                remoteRunspace.Close();
                remoteRunspace.Dispose();
                throw E;
            }
        }

    
        [NonAction]

        public PowerShell RemoveMailBoxFolderPermission(string IdentityName, string UserName, ref Runspace remoteRunspace)
        {
            try
            {
                Command activityCommand = new Command("Remove-MailboxFolderPermission");
                activityCommand.Parameters.Add("-Identity", IdentityName);
                activityCommand.Parameters.Add("-User", UserName);
                activityCommand.Parameters.Add("-Confirm", false);
                PowerShell powershell = PowerShell.Create();
                powershell.Runspace = remoteRunspace;
                powershell.Commands.AddCommand(activityCommand);
                return powershell;
            }
            catch (System.Exception E)
            {
                remoteRunspace.Close();
                remoteRunspace.Dispose();
                throw E;
            }
        }

    }
}