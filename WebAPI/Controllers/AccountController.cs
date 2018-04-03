using Jose;
using WebAPI.Helpers;
using WebAPI.Models;
using System;
using System.Configuration;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Text;
using System.Web.Http;

namespace WebAPI.Controllers
{
    [RoutePrefix("Account")]

    public class AccountController : ApiController
    {
        public AccountController()
        {
            aOpenRunSpace = new OpenRunSpace();
            aAccountHelpers = new AccountHelpers();
            aCommomHelpers = new CommomHelpers();
        }

        private OpenRunSpace aOpenRunSpace;
        private AccountHelpers aAccountHelpers;
        private CommomHelpers aCommomHelpers;

        /// <summary>
        /// 回傳使用者登入授權Token
        /// </summary>
        /// <param name="loginData"></param>
        /// <returns></returns>
        [Route("")]
        public IHttpActionResult Post([FromBody]AccountModel loginData) //UserLoginInfo
        {
            Runspace remoteRunspace = null;
            try
            {
                aOpenRunSpace.Open(loginData.AccountName, loginData.Password, ref remoteRunspace);
                var results = aAccountHelpers.UserData(loginData.AccountName, ref remoteRunspace);
                results.Invoke();
                var ErrorMsgs = aCommomHelpers.ReturnPowerShellInvokeErrors(results.Streams.Error);
                if (ErrorMsgs != null)
                {
                    return Content(HttpStatusCode.BadRequest, ErrorMsgs);
                }
                var secret = ConfigurationManager.AppSettings["JWTKey"].ToString();
                
                var payload = new
                {
                    loginData = loginData
                };

                var Token = Jose.JWT.Encode(payload, Encoding.UTF8.GetBytes(secret), JwsAlgorithm.HS256);
                return Ok(Token);
               
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