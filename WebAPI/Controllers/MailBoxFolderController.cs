using WebAPI.Attributes;
using WebAPI.Helpers;
using Swashbuckle.Swagger.Annotations;
using System;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Xml.Linq;

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