using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Web.Http;
using static WebAPI.Models.MailBoxFolderModels;

namespace WebAPI.Helpers
{
    public class MailBoxFolderHelpers
    {
      
        [NonAction]
        public PowerShell MailBoxFolders(string IdentityName, ref Runspace remoteRunspace)
        {
            PowerShell powershell = PowerShell.Create();
            try
            {
                Command activityCommand = new Command("Get-MailboxFolder");
                activityCommand.Parameters.Add("-Identity", IdentityName);
                activityCommand.Parameters.Add("-MailFolderOnly");
                activityCommand.Parameters.Add("-GetChildren");             
                powershell.Runspace = remoteRunspace;
                powershell.Commands.AddCommand(activityCommand);            
                return powershell;
            }
            catch (Exception E)
            {
                remoteRunspace.Close();
                remoteRunspace.Dispose();
                throw E;
            }
            
        }

       
        [NonAction]

        public IEnumerable<RootFolderModel> RootFolderResults(ICollection<PSObject> aCollection)
        {
            return aCollection.Select(i => new RootFolderModel
            {
                Name = i.Properties["Name"].Value.ToString(),
                HasSubfolders = (Boolean)i.Properties["HasSubfolders"].Value
            });
        }

      
        [NonAction]

        public IEnumerable<SubFolderModel> SubFolderSelectResults(ICollection<PSObject> aCollection)
        {
            return aCollection.Select(i => new SubFolderModel
            {
                Name = i.Properties["Name"].Value.ToString(),
                FolderPath = ((PSObject)(i.Properties["FolderPath"].Value)).ImmediateBaseObject,
                HasSubfolders = (Boolean)i.Properties["HasSubfolders"].Value
            });
        }
    }
}