using Microsoft.AspNetCore.Mvc;
using QPayBackend.Models;
using QPayBackend.Tools;
using System.Threading.Tasks;
using ToolUtilityNameSpace;

namespace QPayBackend.Controllers
{
    [Route("api/[controller]")]
    public class QPayAtmController : Controller
    {
        private ToolUtilityClass m_ToolUtilityClass { get; set; }
        #region 除錯用參數
        private const int TOTAL_LEVEL = 1;//改變這個值，就會改追蹤的階層，值越小越不會追蹤，若是 TOTAL_LEVEL = 3 ，則大於 3 的 LEVEL，例如 : LEVEL_4、LEVEL_5 就不會被追蹤
        //private const int TOTAL_LEVEL = 5;//改變這個值，就會改追蹤的階層，值越大越會追蹤，若是 TOTAL_LEVEL = 3 ，則大於 3 的 LEVEL，例如 : LEVEL_4、LEVEL_5 就不會被追蹤
        private const int LEVEL_1 = 1; // 比較容易被看到的，可能是比較大範圍的部分
        private const int LEVEL_2 = 2;
        private const int LEVEL_3 = 3;
        private const int LEVEL_4 = 4;
        private const int LEVEL_5 = 5; // 比較不會被看到的，可能是比較細節的部分
        // 如果 TRACE_LEVEL >= TRACE_LEVEL_GROUND 就會進行追蹤
        // 如果 TRACE_LEVEL < TRACE_LEVEL_GROUND 就不會進行追蹤
        //int TRACE_LEVEL = 5;
        //int TRACE_LEVEL_GROUND = 3;
        #endregion

        public QPayAtmController()
        {
            //m_ToolUtilityClass = new ToolUtilityClass("DYNAMICS365", "jesus");

            //m_ToolUtilityClass.TraceByLevel(TOTAL_LEVEL, LEVEL_1, "QPayAtmController，Constructor");
        }

        [HttpGet]
        [Route("QPayBackendUrl")]
        public async Task<IActionResult> QPayReturnUrl(int? id = 0)
        {
            //m_ToolUtilityClass = new ToolUtilityClass("DYNAMICS365", "jesus");

            //m_ToolUtilityClass.TraceByLevel(TOTAL_LEVEL, LEVEL_1, "QPayAtmController:QPayBackendUrl-001:ShopNo=XXXXXXXXX");

            return new OkObjectResult("這是永豐金流後台!");
        }

        [HttpPost]
        [Route("QPayBackendUrl")]
        public JsonResult QPayBackendUrl([FromBody] BackendPostData aBackendPostData)
        {
            m_ToolUtilityClass = new ToolUtilityClass("DYNAMICS365", "jesus");

            m_ToolUtilityClass.TraceByLevel(TOTAL_LEVEL, LEVEL_1, "QPayAtmController:QPayBackendUrl-001:ShopNo=" + aBackendPostData.ShopNo);

            using (QPayAtmWebhook aQPayAtmWebhook = new QPayAtmWebhook())
            {
                return aQPayAtmWebhook.QPayBackendUrl(aBackendPostData);
            }
        }
    }
}
