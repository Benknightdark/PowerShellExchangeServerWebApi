using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using WebAPI.Models;

namespace WebAPI.Helpers
{
    public class AzureADHelpers
    {
        public AzureADHelpers()
        {
        }

        public AccountModel UserData;
        public IEnumerable<PSObject> Users(AccountModel model)
        {
            string GetAzureADUsersScript = ConnectAzureAD(model)+ "  Get-AzureADUser";
            return GetPSresults(GetAzureADUsersScript);
        }
        private IEnumerable<PSObject> GetPSresults(string scriptText)
        {
            Runspace runspace = RunspaceFactory.CreateRunspace();

            runspace.Open();

            Pipeline pipeline = runspace.CreatePipeline();

            pipeline.Commands.AddScript(scriptText);

            Collection<PSObject> results = pipeline.Invoke();
            runspace.Close();
            return results.Skip(1);
        }

        private string ConnectAzureAD(AccountModel model)
        {
            string ConnectScript =(string.Format(@"
                                    Import-Module AzureAD
                                    $User ='{0}'
                                    $PWord = ConvertTo-SecureString -String '{1}' -AsPlainText -Force
                                    $Credential = New-Object -TypeName 'System.Management.Automation.PSCredential' -ArgumentList $User, $PWord
                                    Connect-AzureAD -Credential $Credential 
                                   
                ", model.AccountName,model.Password));
            return ConnectScript;
        }
    }
}