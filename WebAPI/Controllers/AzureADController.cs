using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
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

        [HttpGet]
        [Route("Users")]
        [EnableCors(origins: "*", headers: "*", methods: "*")]
        [JWTAttribute]
        public IHttpActionResult GetAzureADusers()
        {
            try
            {
                return Ok(aAzureADHelpers.Users(GetUserData()));
                //
            }
            catch (Exception e)
            {

                return BadRequest(e.Message);
            }
            
        }

    }
}
