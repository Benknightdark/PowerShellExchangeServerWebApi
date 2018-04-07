using System.Collections.ObjectModel;
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
        public Collection<PSObject> Users(AccountModel model)
        {
            string GetAzureADUsersScript = ConnectAzureAD(model) ;
            return GetPSresults(GetAzureADUsersScript);
        }
        public Collection<PSObject> GetPSresults(string scriptText)
        {
            Runspace runspace = RunspaceFactory.CreateRunspace();

            runspace.Open();

            Pipeline pipeline = runspace.CreatePipeline();

            pipeline.Commands.AddScript(scriptText);

            Collection<PSObject> results = pipeline.Invoke();
            runspace.Close();
            return results;
        }

        public string ConnectAzureAD(AccountModel model)
        {
            string ConnectScript =(string.Format(@"
                                    Import-Module AzureAD
                                    $User ='{0}'
                                    $PWord = ConvertTo-SecureString -String '{1}' -AsPlainText -Force
                                    $Credential = New-Object -TypeName 'System.Management.Automation.PSCredential' -ArgumentList $User, $PWord
                                    Connect-AzureAD -Credential $Credential
                                    Get-AzureADUser
                ", model.AccountName,model.Password));
            return ConnectScript;
        }
    }
}