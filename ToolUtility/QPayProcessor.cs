using QPay.Domain;
using System;
using System.Threading.Tasks;
using Line.Messaging;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using ToolUtilityNameSpace;

namespace ToolUtilityNameSpace
{
    public class QPayProcessor
    {
        #region 資料區
        //string m_ShopNo = "NA0149_001";
        string m_ShopNo = "DA1626_001";

        //private const String RETURN_URL = "https://yhchurchback.speechmessage.com.tw:454/api/QPay/QPayReturnUrl";
        //private const String BACKEND_URL = "https://yhchurchback.speechmessage.com.tw:454/api/QPay/QPayBackendUrl";

        private const String RETURN_URL  = "https://yhchurchback.speechmessage.com.tw:454/api/QPayCard/QPayReturnUrl";
        //private const String BACKEND_URL = "https://yhchurchback.speechmessage.com.tw:454/api/QPayAtm/PushSuccess";
        //private const String BACKEND_URL = "http://yhchurchback.speechmessage.com.tw:80/api/QPayAtm/PushSuccess";
        //private const String BACKEND_URL = "http://yhchurchback.speechmessage.com.tw/api/QPayAtm/PushSuccess";
        private const String BACKEND_URL = "http://yhchurchback.speechmessage.com.tw/api/QPayAtm/QPayBackendUrl";

        //private LinePayClient m_LinePayClient { get; }

        #endregion
        #region 初始化
        public QPayProcessor()
        {
        }
        public QPayProcessor(String aShopNo)
        {
            m_ShopNo = aShopNo;
        }
        #endregion
        #region 永豐金流程式區
        public async Task<CreOrder> CreOrderCard(int Amount, String ProductName)
        {
            //設定參數
            CreOrderReq creOrderReq = new CreOrderReq()
            {
                ShopNo = m_ShopNo,
                OrderNo = "C" + DateTime.Now.ToString("yyyyMMddhhmmssfff"),
                Amount = Amount * 100,
                CurrencyID = "TWD",
                PrdtName = ProductName,
                ReturnURL = RETURN_URL,
                BackendURL = BACKEND_URL,
                PayType = "C",
                CardParam = new CreOrderCardParamReq()
                {
                    AutoBilling = "Y"
                }
            };

            CreOrder retObj = QPayToolkit.OrderCreate(creOrderReq);

            var Result = QPayCommon.SerializeToJson(retObj);

            return retObj;

        }
        public async Task<String> CreOrderATM(int Amount, String ProductName)
        {
            //設定參數
            //設定參數
            CreOrderReq creOrderReq = new CreOrderReq()
            {
                ShopNo = m_ShopNo,
                OrderNo = "A" + DateTime.Now.ToString("yyyyMMddhhmmssfff"),
                Amount = Amount * 100,
                CurrencyID = "TWD",
                PrdtName = ProductName,
                ReturnURL = RETURN_URL,
                BackendURL = BACKEND_URL,
                PayType = "A",
                ATMParam = new CreOrderATMParamReq()
                {
                    ExpireDate = DateTime.Now.AddDays(10).ToString("yyyyMMdd")
                }
            };

            CreOrder retObj = QPayToolkit.OrderCreate(creOrderReq);

            //var Result = QPayCommon.SerializeToJson(retObj);

            //String aAtmpayNo =
            //                    "<br/>* 請依照訊息付款 * <br/>" +
            //                    "銀行代碼 : 807 永豐商業銀行<br/>" +
            //                    "分行代號 : 021 台北分行<br/>" +
            //                    "帳號     : " + retObj.ATMParam.AtmPayNo + "<br/>" +
            //                    //"戶名     : 音訊豐富教會<br/>";
            //                    "戶名     : 其他應付款-代收-網路收款<br/>";
            ////return Json(new { status = "1", message = "感謝您的奉獻", QPayUrl = retObj.ATMParam.OtpURL });
            ////return Json(new { status = "2", message = "請依照訊息指示付款", AtmpayNo = aAtmpayNo });
            //return retObj;

            //return "<br/>* 請依照訊息付款 * <br/>" +
            //                    "銀行代碼 : 807 永豐商業銀行<br/>" +
            //                    "分行代號 : 021 台北分行<br/>" +
            //                    "帳號     : " + retObj.ATMParam.AtmPayNo + "<br/>" +
            //                    //"戶名     : 音訊豐富教會<br/>";
            //                    "戶名     : 其他應付款-代收-網路收款<br/>";
            return
                "* 請依照訊息付款 *" + Environment.NewLine +
                "銀行代碼 : 807 永豐商業銀行" + Environment.NewLine +
                "分行代號 : 021 台北分行" + Environment.NewLine +
                "帳號     : " + retObj.ATMParam.AtmPayNo + Environment.NewLine +
                //"戶名     : 音訊豐富教會<br/>";
                "戶名     : 其他應付款-代收-網路收款";

            //String WebAtmUrl = retObj.ATMParam.WebAtmURL.TrimEnd(new Char[] { ' ', '\\', '.' });
            //return Json(new { status = "1", message = "感謝您的奉獻", QPayUrl = WebAtmUrl });

        }
        public async Task<CreOrder> CreOrderCard(int Amount, String ProductName, String OrderDate)
        {
            //設定參數
            CreOrderReq creOrderReq = new CreOrderReq()
            {
                ShopNo = m_ShopNo,
                OrderNo = "C" + OrderDate,
                Amount = Amount * 100,
                CurrencyID = "TWD",
                PrdtName = ProductName,
                ReturnURL = RETURN_URL,
                BackendURL = BACKEND_URL,
                PayType = "C",
                CardParam = new CreOrderCardParamReq()
                {
                    AutoBilling = "Y"
                }
            };

            CreOrder retObj = QPayToolkit.OrderCreate(creOrderReq);

            var Result = QPayCommon.SerializeToJson(retObj);

            return retObj;

        }
        public async Task<String> CreOrderATM(int Amount, String ProductName, String OrderDate)
        {
            //設定參數
            //設定參數
            CreOrderReq creOrderReq = new CreOrderReq()
            {
                ShopNo = m_ShopNo,
                OrderNo = "A" + OrderDate,
                Amount = Amount * 100,
                CurrencyID = "TWD",
                PrdtName = ProductName,
                ReturnURL = RETURN_URL,
                BackendURL = BACKEND_URL,
                PayType = "A",
                ATMParam = new CreOrderATMParamReq()
                {
                    ExpireDate = DateTime.Now.AddDays(10).ToString("yyyyMMdd")
                }
            };

            CreOrder retObj = QPayToolkit.OrderCreate(creOrderReq);

            //var Result = QPayCommon.SerializeToJson(retObj);

            //String aAtmpayNo =
            //                    "<br/>* 請依照訊息付款 * <br/>" +
            //                    "銀行代碼 : 807 永豐商業銀行<br/>" +
            //                    "分行代號 : 021 台北分行<br/>" +
            //                    "帳號     : " + retObj.ATMParam.AtmPayNo + "<br/>" +
            //                    //"戶名     : 音訊豐富教會<br/>";
            //                    "戶名     : 其他應付款-代收-網路收款<br/>";
            ////return Json(new { status = "1", message = "感謝您的奉獻", QPayUrl = retObj.ATMParam.OtpURL });
            ////return Json(new { status = "2", message = "請依照訊息指示付款", AtmpayNo = aAtmpayNo });
            //return retObj;

            //return "<br/>* 請依照訊息付款 * <br/>" +
            //                    "銀行代碼 : 807 永豐商業銀行<br/>" +
            //                    "分行代號 : 021 台北分行<br/>" +
            //                    "帳號     : " + retObj.ATMParam.AtmPayNo + "<br/>" +
            //                    //"戶名     : 音訊豐富教會<br/>";
            //                    "戶名     : 其他應付款-代收-網路收款<br/>";
            return
                "* 請依照訊息付款 *" + Environment.NewLine +
                "銀行代碼 : 807 永豐商業銀行" + Environment.NewLine +
                "分行代號 : 021 台北分行" + Environment.NewLine +
                "帳號     : " + retObj.ATMParam.AtmPayNo + Environment.NewLine +
                //"戶名     : 音訊豐富教會<br/>";
                "戶名     : 其他應付款-代收-網路收款";

            //String WebAtmUrl = retObj.ATMParam.WebAtmURL.TrimEnd(new Char[] { ' ', '\\', '.' });
            //return Json(new { status = "1", message = "感謝您的奉獻", QPayUrl = WebAtmUrl });

        }
        public async Task<CreOrder> CreateOrderATM(int Amount, String ProductName)
        {
            //設定參數
            //設定參數
            CreOrderReq creOrderReq = new CreOrderReq()
            {
                ShopNo = m_ShopNo,
                OrderNo = "A" + DateTime.Now.ToString("yyyyMMddhhmmssfff"),
                Amount = Amount * 100,
                CurrencyID = "TWD",
                PrdtName = ProductName,
                ReturnURL = RETURN_URL,
                BackendURL = BACKEND_URL,
                PayType = "A",
                ATMParam = new CreOrderATMParamReq()
                {
                    ExpireDate = DateTime.Now.AddDays(10).ToString("yyyyMMdd")
                }
            };

            return QPayToolkit.OrderCreate(creOrderReq);

        }
        public async Task<CreOrder> CreateOrderATM(int Amount, String ProductName, String OrderDate)
        {
            //設定參數
            //設定參數
            CreOrderReq creOrderReq = new CreOrderReq()
            {
                ShopNo = m_ShopNo,
                OrderNo = "A" + OrderDate,
                Amount = Amount * 100,
                CurrencyID = "TWD",
                PrdtName = ProductName,
                ReturnURL = RETURN_URL,
                BackendURL = BACKEND_URL,
                PayType = "A",
                ATMParam = new CreOrderATMParamReq()
                {
                    ExpireDate = DateTime.Now.AddDays(10).ToString("yyyyMMdd")
                }
            };

            return QPayToolkit.OrderCreate(creOrderReq);

        }
        public async Task<QryOrder> OrderQuery(String aOrderNo)
        {
            QryOrderReq orderQueryReq = new QryOrderReq()
            {
                ShopNo = m_ShopNo,
                OrderNo = aOrderNo
            };

            QryOrder retObj = QPayToolkit.OrderQuery(orderQueryReq);

            return retObj;
        }
        public QryOrderPay OrderPayQuery(String aPayToken)
        {
            QryOrderPayReq orderPayQueryReq = new QryOrderPayReq()
            {
                ShopNo = m_ShopNo,
                PayToken = aPayToken
            };

            QryOrderPay retObj = QPayToolkit.OrderPayQuery(orderPayQueryReq, ConvertShopNoToHashCode(m_ShopNo));

            return retObj;
        }
        public QryOrderPay OrderPayQuery( String aShopNo, String aPayToken)
        {
            QryOrderPayReq orderPayQueryReq = new QryOrderPayReq()
            {
                ShopNo = aShopNo,
                PayToken = aPayToken
            };

            QryOrderPay retObj = QPayToolkit.OrderPayQuery(orderPayQueryReq, ConvertShopNoToHashCode(aShopNo));

            return retObj;
        }
        public QryOrder OrderQuery(String aShopNo, String OrderNo)
        {
            QryOrderReq orderQueryReq = new QryOrderReq()
            {
                ShopNo = aShopNo,
                OrderNo = OrderNo
            };

            QryOrder retObj = QPayToolkit.OrderQuery(orderQueryReq);

            return retObj;
        }
        public async Task<QryBill> BillQuery(String aPayDate)
        {
            QryBillReq billQueryReq = new QryBillReq()
            {
                ShopNo = m_ShopNo,
                BillDate = aPayDate
            };

            QryBill retObj = QPayToolkit.BillQuery(billQueryReq);

            //ltResponse.Text = QPayCommon.SerializeToJson(retObj);
            return retObj;
        }
        public async Task<QryAllot> AllotQuery(String aAllotDateS, String aAllotDateE, String aPayType)
        {
            QryAllotReq allotQueryReq = new QryAllotReq()
            {
                ShopNo = m_ShopNo,
                AllotDateS = aAllotDateS,
                AllotDateE = aAllotDateE,
                PayType = aPayType
            };

            QryAllot retObj = QPayToolkit.AllotQuery(allotQueryReq);

            //ltResponse.Text = QPayCommon.SerializeToJson(retObj);
            return retObj;
        }
        public async Task<QryOrderUnCaptured> OrderUnCapturedQuery()
        {
            QryOrderUnCapturedReq orderUnCapturedReq = new QryOrderUnCapturedReq()
            {
                ShopNo = m_ShopNo
            };

            QryOrderUnCaptured retObj = QPayToolkit.OrderUnCapturedQuery(orderUnCapturedReq);

            //ltResponse.Text = QPayCommon.SerializeToJson(retObj);
            return retObj;
        }
        public async Task<OrderMaintain> OrderMaintain(String aOrderNo, String aCommand)
        {
            OrderMaintainReq orderMaintainReq = new OrderMaintainReq()
            {
                ShopNo = m_ShopNo,
                OrderNo = aOrderNo,
                Command = aCommand
            };

            OrderMaintain retObj = QPayToolkit.OrderMaintain(orderMaintainReq);

            //ltResponse.Text = QPayCommon.SerializeToJson(retObj);
            return retObj;
        }
        #endregion
        #region 永豐金流工具區
        private string ConvertShopNoToHashCode(String aShopNo)
        {
            //客製化
            switch (aShopNo)
            {
                case "DA1626_001":
                    // 永和禮拜堂(公司研發)
                    return "D1695F439A69448F,7E460E920A184845,DEA83EFB714943F3,DC237C5C69914F0C";
                case "NA0149_001":
                    // 音訊教會 SandBox
                    return "5E854757C751413F,D743D0EB06904837,08169D5445644513,8E52B5A180EE4399";
                default:
                    return "5E854757C751413F,D743D0EB06904837,08169D5445644513,8E52B5A180EE4399";
            }
        }
        #endregion
    }

}
