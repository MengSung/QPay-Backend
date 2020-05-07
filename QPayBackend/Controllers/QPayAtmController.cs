using Line.Messaging;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Xrm.Sdk;
using QPay.Domain;
using System;
using System.Threading.Tasks;
using ToolUtilityNameSpace;
using System.Web;
using QPayBackend.Models;
using System.Collections.Generic;
using QPayBackend.Tools;

namespace QPayBackend.Controllers
{
    [Route("api/[controller]")]
    public class QPayAtmController : Controller
    {
        public QPayAtmController()
        {
        }

        [HttpGet]
        [Route("QPayBackendUrl")]
        public async Task<IActionResult> QPayReturnUrl(int? id = 0)
        {
            return new OkObjectResult("付款結果可能成功");
        }

        [HttpPost]
        [Route("QPayBackendUrl")]
        public JsonResult QPayBackendUrl([FromBody] BackendPostData aBackendPostData)
        {
            using (QPayAtmWebhook aQPayAtmWebhook = new QPayAtmWebhook())
            {
                return aQPayAtmWebhook.QPayBackendUrl(aBackendPostData);
            }
            //QPayAtmWebhook aQPayAtmWebhook = new QPayAtmWebhook();
            //return aQPayAtmWebhook.QPayBackendUrl(aBackendPostData);
        }
    }
}
