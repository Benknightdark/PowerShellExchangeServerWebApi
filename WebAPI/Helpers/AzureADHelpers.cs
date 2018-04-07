using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Web;

namespace WebAPI.Helpers
{
    public class AzureADHelpers
    {
        public AzureADHelpers(dynamic _UserData)
        {
            UserData = _UserData;
        }
        public dynamic UserData;
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
    }
}