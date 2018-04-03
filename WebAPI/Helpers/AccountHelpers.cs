using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Web;

namespace WebAPI.Helpers
{
    public class AccountHelpers
    {
        public PowerShell UserData(string IdentityName, ref Runspace remoteRunspace)
        {
            PowerShell powershell = PowerShell.Create();
            try
            {
                Command activityCommand = new Command("Get-User");
                activityCommand.Parameters.Add("-Identity", IdentityName);
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
    }
}