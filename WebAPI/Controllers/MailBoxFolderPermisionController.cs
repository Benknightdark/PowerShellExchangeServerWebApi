using WebAPI.Attributes;
using WebAPI.Helpers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Web.Http;
using System.Web.Http.Cors;
using static WebAPI.Models.MailBoxFolderPermissionModels;

namespace WebAPI.Controllers
{
    [RoutePrefix("MailBoxFolderPermission")]
    public class MailBoxFolderPermisionController : BaseController
    {
        public MailBoxFolderPermisionController()
        {
            aOpenRunSpace = new OpenRunSpace();
            aMailBoxFolderPermissionHelpers = new MailBoxFolderPermissionHelpers();
            aMailBoxFolderHelpers = new MailBoxFolderHelpers();
            aCommomHelpers = new CommomHelpers();
        }

        private OpenRunSpace aOpenRunSpace;
        private MailBoxFolderPermissionHelpers aMailBoxFolderPermissionHelpers;
        private MailBoxFolderHelpers aMailBoxFolderHelpers;
        private CommomHelpers aCommomHelpers;

        /// <summary>
        /// 資料夾權限
        /// </summary>
        /// <param name="Name"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("")]
        [EnableCors(origins: "*", headers: "*", methods: "*")]
        [JWTAttribute]
        public IHttpActionResult GetMailBoxFolderPermission([FromUri]string Name)
        {
            Runspace remoteRunspace = null;
            try
            {
                aOpenRunSpace.Open(GetUserData().AccountName, GetUserData().Password, ref remoteRunspace);
                var results = aMailBoxFolderPermissionHelpers.PSmailBoxFolderPermission(GetUserData().AccountName + ":\\" + Name, ref remoteRunspace);
                var RootFolderPermissionResults = aMailBoxFolderPermissionHelpers.RootFolderPermissionResults(results.Invoke());
                var ErrorMsgs = aCommomHelpers.ReturnPowerShellInvokeErrors(results.Streams.Error);
                if (ErrorMsgs != null)
                {
                    return Content(HttpStatusCode.BadRequest, ErrorMsgs);
                }
                return Ok(RootFolderPermissionResults);
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

        /// <summary>
        /// 新增資料夾權限
        /// </summary>
        /// <param name="Model"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("")]
        [EnableCors(origins: "*", headers: "*", methods: "*")]
        [JWTAttribute]
        public IHttpActionResult PostMailBoxFolderPermission([FromBody] List<PostPutMailBoxFolderPermission> Model)
        {
            Runspace remoteRunspace = null;
            try
            {
                aOpenRunSpace.Open(GetUserData().AccountName, GetUserData().Password, ref remoteRunspace);
                foreach (var m in Model)
                {
                    if (m.AccessRights.Length > 0)
                    {
                        var results = aMailBoxFolderPermissionHelpers.AddMailBoxFolderPermission(GetUserData().AccountName + ":\\" + m.FolderPath, m.UserPrincipalName, m.AccessRights, ref remoteRunspace);
                        results.Invoke();
                        var ErrorMsgs = aCommomHelpers.ReturnPowerShellInvokeErrors(results.Streams.Error);
                        if (ErrorMsgs != null)
                        {
                            return Content(HttpStatusCode.BadRequest, ErrorMsgs);
                        }
                        results.Commands.Clear();
                    }
                }
                RecursiveSettingAccessRights(Model, ref remoteRunspace);

                return Ok();
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

        /// <summary>
        /// 修改資料夾權限
        /// </summary>
        /// <param name="Model"></param>
        /// <returns></returns>
        [HttpPut]
        [Route("")]
        [EnableCors(origins: "*", headers: "*", methods: "*")]
        [JWTAttribute]
        public IHttpActionResult PutMailBoxFolderPermission([FromBody] List<PostPutMailBoxFolderPermission> Model)
        {
            Runspace remoteRunspace = null;
            try
            {
                aOpenRunSpace.Open(GetUserData().AccountName, GetUserData().Password, ref remoteRunspace);
                foreach (var m in Model)
                {
                    if (m.AccessRights.Length > 0)
                    {
                        var results = aMailBoxFolderPermissionHelpers.SetMailBoxFolderPermission(GetUserData().AccountName + ":\\" + m.FolderPath, m.UserPrincipalName, m.AccessRights, ref remoteRunspace);
                        results.Invoke();
                        var ErrorMsgs = aCommomHelpers.ReturnPowerShellInvokeErrors(results.Streams.Error);
                        if (ErrorMsgs != null)
                        {
                            return Content(HttpStatusCode.BadRequest, ErrorMsgs);
                        }
                        results.Commands.Clear();
                    }
                }
                RecursiveSettingAccessRights(Model, ref remoteRunspace);
                return Ok();
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

        /// <summary>
        /// 刪除資料夾權限
        /// </summary>
        /// <param name="Name"></param>
        /// <param name="FolderPath"></param>
        /// <returns></returns>
        [HttpDelete]
        [Route("")]
        [EnableCors(origins: "*", headers: "*", methods: "*")]
        [JWTAttribute]
        public IHttpActionResult DeleteMailBoxFolderPermission([FromUri]string Name, string FolderPath)
        {
            Runspace remoteRunspace = null;
            try
            {
                aOpenRunSpace.Open(GetUserData().AccountName, GetUserData().Password, ref remoteRunspace);
                var results = aMailBoxFolderPermissionHelpers.RemoveMailBoxFolderPermission(GetUserData().AccountName + ":\\" + FolderPath, Name, ref remoteRunspace);
                results.Invoke();
                var ErrorMsgs = aCommomHelpers.ReturnPowerShellInvokeErrors(results.Streams.Error);
                if (ErrorMsgs != null)
                {
                    return Content(HttpStatusCode.BadRequest, ErrorMsgs);
                }
                results.Commands.Clear();

                return Ok();
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
        
        [NonAction]
        private void RecursiveSettingAccessRights(List<PostPutMailBoxFolderPermission> Model, ref Runspace remoteRunspace)
        {
            var SubFolderResults = aMailBoxFolderHelpers.MailBoxFolders(GetUserData().AccountName + ":\\" + Model.FirstOrDefault().FolderPath, ref remoteRunspace);
            var SubFolderSelectResults = aMailBoxFolderHelpers.SubFolderSelectResults(SubFolderResults.Invoke());
            foreach (var sub in SubFolderSelectResults)
            {
                string aFolderPath = string.Join("\\", ((ArrayList)sub.FolderPath).Cast<object>().Select(x => x == null ? null : x.ToString()).ToArray()).ToString();
                //加入權限
                foreach (var m in Model)
                {
                    if (m.AccessRights.Length > 0)
                    {
                        string AccountName = GetUserData().AccountName;
                        var MailBoxFolderPermissionResults = aMailBoxFolderPermissionHelpers.PSmailBoxFolderPermission(AccountName + ":\\" + aFolderPath, ref remoteRunspace);
                        var MailBoxFolderPermissionSelectResults = aMailBoxFolderPermissionHelpers.RootFolderPermissionResults(MailBoxFolderPermissionResults.Invoke());                      
                        PowerShell results = null;
                        if (MailBoxFolderPermissionSelectResults.Where(a => a.UserPrincipalName == m.UserPrincipalName).Any())
                        {
                            results = aMailBoxFolderPermissionHelpers.SetMailBoxFolderPermission(GetUserData().AccountName+":\\" + aFolderPath, m.UserPrincipalName, m.AccessRights, ref remoteRunspace);
                        }
                        else
                        {
                            results = aMailBoxFolderPermissionHelpers.AddMailBoxFolderPermission(GetUserData().AccountName+":\\" + aFolderPath, m.UserPrincipalName, m.AccessRights, ref remoteRunspace);
                        }
                        results.Invoke();
                        //var ErrorMsgs = aCommomHelpers.ReturnPowerShellInvokeErrors(results.Streams.Error);
                        //if (ErrorMsgs != null)
                        //{
                        //    remoteRunspace.Close();
                        //    remoteRunspace.Dispose();
                        //    return;
                        //}
                        results.Commands.Clear();
                    }
                }
                //遞迴查詢
                var aModel = Model;
                foreach (var a in aModel.ToList())
                {
                    a.FolderPath = aFolderPath;
                }
                if (sub.HasSubfolders)
                {
                    RecursiveSettingAccessRights(aModel, ref remoteRunspace);
                }
            }
        }
    }
}