using Jose;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System.Configuration;
using System.Dynamic;
using System.Text;
using System.Web.Http;
using WebAPI.Models;

namespace WebAPI.Controllers
{
    public class BaseController : ApiController
    {
        public BaseController()
        {
        }
        [NonAction]
        public  AccountModel GetUserData()
        {
            var secret = ConfigurationManager.AppSettings["JWTKey"].ToString();
            var jwtObject = Jose.JWT.Decode(Request.Headers.Authorization.Parameter,Encoding.UTF8.GetBytes(secret),JwsAlgorithm.HS256);
            dynamic obj = JsonConvert.DeserializeObject<ExpandoObject>(jwtObject, new ExpandoObjectConverter());
            AccountModel m = new AccountModel();
            m.AccountName = obj.loginData.AccountName;
            m.Password = obj.loginData.Password;
            return m;
        }
    }
}