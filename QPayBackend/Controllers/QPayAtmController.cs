using Microsoft.AspNetCore.Mvc;
using QPayBackend.Models;
using QPayBackend.Tools;
using System.Threading.Tasks;

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
            return new OkObjectResult("這是永豐金流後台!");
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
