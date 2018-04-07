using WebAPI.Attributes;
using WebAPI.Helpers;
using System;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Xml.Linq;

namespace WebAPI.Controllers
{
    [RoutePrefix("SubMailBoxFolder")]
    public class SubMailBoxFolderController : BaseController
    {
        public SubMailBoxFolderController()
        {
            aOpenRunSpace = new OpenRunSpace();
            aMailBoxFolderHelpers = new MailBoxFolderHelpers();
            aCommomHelpers = new CommomHelpers();
        }

        private OpenRunSpace aOpenRunSpace;
        private MailBoxFolderHelpers aMailBoxFolderHelpers;
        private CommomHelpers aCommomHelpers;

       

        /// <summary>
        /// 子郵件資料夾
        /// </summary>
        /// <param name="Name"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("")]
        [EnableCors(origins: "*", headers: "*", methods: "*")]
        [JWTAttribute]
        public IHttpActionResult GetSubMailBoxFolder([FromUri]string Name)
        {
            Runspace remoteRunspace = null;
            try
            {
                aOpenRunSpace.Open(GetUserData().AccountName, GetUserData().Password, ref remoteRunspace);
                var results = aMailBoxFolderHelpers.MailBoxFolders(GetUserData().AccountName + ":\\" + Name, ref remoteRunspace);
                var SubFolderSelectResults = aMailBoxFolderHelpers.SubFolderSelectResults(results.Invoke());
                var ErrorMsgs = aCommomHelpers.ReturnPowerShellInvokeErrors(results.Streams.Error);
                if (ErrorMsgs != null)
                {
                    return Content(HttpStatusCode.BadRequest, ErrorMsgs);
                }
                return Ok(SubFolderSelectResults);
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
                remoteRunspace.Dispose();
            }
        }
    }
}