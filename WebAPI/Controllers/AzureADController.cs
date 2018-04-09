using System;
using System.Linq;
using System.Web.Http;
using System.Web.Http.Cors;
using WebAPI.Attributes;
using WebAPI.Helpers;
using WebAPI.Models;

namespace WebAPI.Controllers
{
    [RoutePrefix("AzureAD")]
    public class AzureADController : BaseController
    {
        public AzureADController()
        {
            aCommomHelpers = new CommomHelpers();
            aAzureADHelpers = new AzureADHelpers();
        }
        public CommomHelpers aCommomHelpers;
        public AzureADHelpers aAzureADHelpers;

        /// <summary>
        /// 回傳azure ad 所有使用者
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("Users")]
        [EnableCors(origins: "*", headers: "*", methods: "*")]
        [JWTAttribute]
        public IHttpActionResult GetAzureADusers()
        {
            try
            {
                var AzureADUserData = aAzureADHelpers.AzureADUsers(GetUserData())
                    .Select(a => new
                    {
                        DisplayName = a.Properties["DisplayName"].Value.ToString(),
                        UserPrincipalName = a.Properties["UserPrincipalName"].Value.ToString()
                    });
                return Ok(AzureADUserData);
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);
            }
        }
    }
}