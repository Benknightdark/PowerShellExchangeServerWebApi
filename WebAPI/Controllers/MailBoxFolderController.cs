using WebAPI.Attributes;
using WebAPI.Helpers;
using System;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Xml.Linq;
using System.Management.Automation;
using System.Collections.ObjectModel;

namespace WebAPI.Controllers
{
    [RoutePrefix("MailBoxFolder")]
    public class MailBoxFolderController : BaseController
    {
        public MailBoxFolderController()
        {
            aOpenRunSpace = new OpenRunSpace();
            aMailBoxFolderHelpers = new MailBoxFolderHelpers();
            aCommomHelpers = new CommomHelpers();
        }

        private OpenRunSpace aOpenRunSpace;
        private MailBoxFolderHelpers aMailBoxFolderHelpers;
        private CommomHelpers aCommomHelpers;
        [HttpGet]
        [Route("userlist")]
        public IHttpActionResult GetUserList()
        {
            using (PowerShell PowerShellInstance = PowerShell.Create())
            {
                // use "AddScript" to add the contents of a script file to the end of the execution pipeline.
                // use "AddCommand" to add individual commands/cmdlets to the end of the execution pipeline.
                PowerShellInstance.AddScript("param($param1) $d = get-date; $s = 'test string value'; " +
                        "$d; $s; $param1; get-service");

                // use "AddParameter" to add a single parameter to the last command/script on the pipeline.
                PowerShellInstance.AddParameter("param1", "parameter 1 value!");
                Collection<PSObject> PSOutput = PowerShellInstance.Invoke();
                return Ok(PSOutput);
            }

        }
        /// <summary>
        /// 郵件資料夾
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("")]
        [EnableCors(origins: "*", headers: "*", methods: "*")]
        [JWTAttribute]
        public IHttpActionResult GetMailBoxFolder()
        {
            Runspace remoteRunspace = null;
            try
            {
                aOpenRunSpace.Open(GetUserData().AccountName, GetUserData().Password, ref remoteRunspace);
                var results = aMailBoxFolderHelpers.MailBoxFolders(GetUserData().AccountName, ref remoteRunspace);
                var RootFolderResults = aMailBoxFolderHelpers.RootFolderResults(results.Invoke());
                var ErrorMsgs = aCommomHelpers.ReturnPowerShellInvokeErrors(results.Streams.Error);
                if (ErrorMsgs != null)
                {
                    return Content(HttpStatusCode.BadRequest, ErrorMsgs);
                }
                return Ok(RootFolderResults);
            }
            catch (Exception e)
            {
                if (e.InnerException != null)
                {
                    return BadRequest(e.InnerException.Message);
                }
                else
                {
                    return BadRequest(e.Message);
                }
            }
            finally
            {
                remoteRunspace.Close();
            }
        }
       

    }
}