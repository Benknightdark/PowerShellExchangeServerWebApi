using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Web;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace WebAPI.Helpers
{
    public class OpenRunSpace
    {
        public void Open(string username, string livePass, ref Runspace remoteRunspace)
        {
            try
            {            
                string uri = "https://outlook.office365.com/powershell-liveid/";
                string schema = "http://schemas.microsoft.com/powershell/Microsoft.Exchange";
                SecureString password = new SecureString();
                foreach (char c in livePass.ToCharArray())
                {
                    password.AppendChar(c);
                }
                PSCredential psc = new PSCredential(username, password);
                WSManConnectionInfo rri = new WSManConnectionInfo(new Uri(uri), schema, psc);
                rri.AuthenticationMechanism = AuthenticationMechanism.Basic;
                remoteRunspace = RunspaceFactory.CreateRunspace(rri);
                remoteRunspace.Open();       
            }
            catch (Exception e)
            {

                remoteRunspace.Close();
                remoteRunspace.Dispose();
                throw e;
            }
            
        }
    }
}