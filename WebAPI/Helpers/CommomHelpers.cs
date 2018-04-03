using WebAPI.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Web;

namespace WebAPI.Helpers
{
    public class CommomHelpers
    {
        public List<PowerShellInvokErrorModel> ReturnPowerShellInvokeErrors(PSDataCollection<ErrorRecord>  ResultsStreamsError)
        {
            PSDataCollection<ErrorRecord> errors = ResultsStreamsError;// results.Streams.Error;
            if (ResultsStreamsError != null && ResultsStreamsError.Count > 0)
            {
                List<PowerShellInvokErrorModel> ErrorMsgs = new List<PowerShellInvokErrorModel>();

                foreach (ErrorRecord err in ResultsStreamsError)
                {
                    ErrorMsgs.Add(new PowerShellInvokErrorModel()
                    {
                        ErrorMsg = err.ToString()
                    });
                }
                return ErrorMsgs;//Content(HttpStatusCode.BadRequest, ErrorMsgs);
            }
            else
            {
                return null;
            }
        }
    }
}