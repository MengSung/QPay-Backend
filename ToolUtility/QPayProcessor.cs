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
        string m_ShopNo = "NA0149_001";

        //private const String RETURN_URL = "https://yhchurchback.speechmessage.com.tw:454/api/QPay/QPayReturnUrl";
        //private const String BACKEND_URL = "https://yhchurchback.speechmessage.com.tw:454/api/QPay/QPayBackendUrl";

        private const String RETURN_URL  = "https://yhchurchback.speechmessage.com.tw:454/api/QPayCard/QPayReturnUrl";
        //private const String BACKEND_URL = "https://yhchurchback.speechmessage.com.tw:454/api/QPayAtm/PushSuccess";
        //private const String BACKEND_URL = "http://yhchurchback.speechmessage.com.tw:80/api/QPayAtm/PushSuccess";
        //private const String BACKEND_URL = "http://yhchurchback.speechmessage.com.tw/api/QPayAtm/PushSuccess";
        private const String BACKEND_URL = "http://yhchurchback.speechmessage.com.tw/api/QPayAtm/QPayBackendUrl";

        //private LinePayClient m_LinePayClient { get; }

        ToolUtilityClass m_ToolUtilityClass = new ToolUtilityClass("DYNAMICS365");
        private LineMessagingClient m_LineMessagingClient { get; }
        private PushUtility m_PushUtility { get; }
        private ReplyUtility m_ReplyUtility { get; }

        #endregion
        #region 初始化
        public QPayProcessor(LineMessagingClient aLineMessagingClient, PushUtility aPushUtility, ReplyUtility aReplyUtility)
        {
            m_LineMessagingClient = aLineMessagingClient;
            //m_LinePayClient = LinePayClient;

            m_PushUtility = aPushUtility;
            m_ReplyUtility = aReplyUtility;
        }
        #endregion
        #region Line 顯示區
        public async Task NotifyQPay(Guid NewStorLessonId, String DisplayLineId, string UserId, Entity aLessonEntity, String ReplyToken)
        {
            try
            {
                #region 通知住綁定的輸入格式
                Entity aContact = this.m_ToolUtilityClass.RetrieveContactCollectionByLineId(UserId);

                String DisplayName = this.m_ToolUtilityClass.GetEntityStringAttribute(ref aContact, "lastname");

                String ProductName = this.m_ToolUtilityClass.GetEntityStringAttribute(aLessonEntity, "new_name");

                // 課程圖案畫面
                String ProductImageUrl = this.m_ToolUtilityClass.GetEntityStringAttribute(aLessonEntity, "new_picture_link");

                // 如果課程沒填圖案畫面
                ProductImageUrl = ProductImageUrl != "" ? ProductImageUrl : "https://web.opendrive.com/api/v1/download/file.json/ODdfMzkyNzk4MV8?inline=1";

                int Amount = (int)this.m_ToolUtilityClass.GetEntityMoneyAttribute(aLessonEntity, "new_lessons_fee").Value;

                //String NotifiedMessage = DisplayName.Replace("(Line)", "") + "已報名" + ProductName + Environment.NewLine +
                //                         "費用 : " + Amount.ToString() + Environment.NewLine +
                //                         "請點選繳費，謝謝您!";
                String NotifiedMessage = ProductName + Environment.NewLine +
                                         "費用 : " + Amount.ToString() + "元" + Environment.NewLine +
                                         "請點選繳費，謝謝您!";

                String OrderDate = DateTime.Now.ToString("yyyyMMddhhmmssfff");

                CreOrder CreatedCardOrder = await CreOrderCard(Amount, ProductName, OrderDate);

                CreOrder CreatedAtmOrder = await CreateOrderATM(Amount, ProductName, OrderDate);

                String AtmResult =
                        "* 請依照訊息付款 *" + Environment.NewLine +
                        "銀行代碼 : 807 永豐商業銀行" + Environment.NewLine +
                        "分行代號 : 021 台北分行" + Environment.NewLine +
                        "帳號     : " + CreatedAtmOrder.ATMParam.AtmPayNo + Environment.NewLine +
                        //"戶名     : 音訊豐富教會<br/>";
                        "戶名     : 其他應付款-代收-網路收款";
                // 建立收費單
                CreateFee( UserId, aLessonEntity, NewStorLessonId , CreatedCardOrder, CreatedAtmOrder );

                List<ITemplateAction> BindingAction = new List<ITemplateAction>();

                BindingAction.Add(new UriTemplateAction("信用卡繳費", CreatedCardOrder.CardParam.CardPayURL));
                BindingAction.Add(new MessageTemplateAction("ATM轉帳/匯款", AtmResult));

                String ButtonMessage = ProductName + Environment.NewLine +
                         "費用 : " + Amount.ToString() + "元";

                //await m_PushUtility.PostSerializedTemplate(DisplayLineId, "有一個繳費的動作必須在手機上操作才能完成", ProductImageUrl, ProductName, "費用 : " + Amount.ToString() + "元", BindingAction);
                await m_ReplyUtility.PostSerializedTemplate(ReplyToken, "有一個繳費的動作必須在手機上操作才能完成", ProductImageUrl, ProductName, "費用 : " + Amount.ToString() + "元", BindingAction);
                //await m_PushUtility.SendMessage(DisplayLineId, "付款網址 = " + CreOrder.CardParam.CardPayURL);

                #endregion

            }
            catch (System.Exception e)
            {
                String ErrorString = "ERROR : FullName = " + this.GetType().FullName.ToString() + " , Time = " + DateTime.Now.ToString() + " , Description = " + e.ToString();

                //Monitor.Exit(this);
                throw e;
            }
        }
        public void CreateFee(string UserId, Entity aLessonEntity, Guid NewStorLessonId, CreOrder CreatedCardOrder, CreOrder CreatedAtmOrder)
        {
            try
            {
                #region 通知住綁定的輸入格式

                Entity aFeeToCreated = new Entity("new_fee");

                SetFeeParameter(aFeeToCreated, UserId, aLessonEntity, NewStorLessonId, CreatedCardOrder, CreatedAtmOrder );

                this.m_ToolUtilityClass.CreateEntity(aFeeToCreated);
                #endregion
            }
            catch (System.Exception e)
            {
                String ErrorString = "ERROR : FullName = " + this.GetType().FullName.ToString() + " , Time = " + DateTime.Now.ToString() + " , Description = " + e.ToString();

                //Monitor.Exit(this);
                throw e;
            }
        }
        public void SetFeeParameter(Entity aFeeToCreated, string UserId, Entity aLessonEntity, Guid NewStorLessonId, CreOrder CreatedCardOrder, CreOrder CreatedAtmOrder)
        {
            try
            {
                #region 通知住綁定的輸入格式
                // 連絡人姓名
                Entity aContact = this.m_ToolUtilityClass.RetrieveContactCollectionByLineId(UserId);

                // 取得課程名稱
                String LessonDisplayName = this.m_ToolUtilityClass.GetEntityStringAttribute(ref aLessonEntity, "new_name");

                // 取得報名者的全名
                String FullName = "";
                FullName = this.m_ToolUtilityClass.GetEntityStringAttribute(ref aContact, "fullname");

                // 收費單名稱
                this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeToCreated, "new_name", LessonDisplayName + "_" + FullName);

                // 收費單姓名關聯 LOOKUP
                this.m_ToolUtilityClass.SetEntityLookUpAttribute(ref aFeeToCreated, "new_contact_new_fee", "contact", aContact.Id);

                // 收費單課程關聯LOOKUP
                this.m_ToolUtilityClass.SetEntityLookUpAttribute(ref aFeeToCreated, "new_disciple_lessons_new_fee", "new_disciple_lessons", aLessonEntity.Id);

                // 收費單應收金額
                this.m_ToolUtilityClass.SetEntityMoneyAttribute(ref aFeeToCreated, "new_fee_shoud_pay", this.m_ToolUtilityClass.GetEntityMoneyAttribute(aLessonEntity, "new_lessons_fee"));
                //this.m_ToolUtilityClass.SetEntityMoneyAttribute(ref aFeeToCreated, "new_fee_really_paid", this.m_ToolUtilityClass.GetEntityMoneyAttribute(aLessonEntity, "new_lessons_fee"));

                // 收費單實收金額
                this.m_ToolUtilityClass.SetEntityMoneyAttribute(ref aFeeToCreated, "new_fee_really_paid", new Money(0));

                // 收費單付款方式
                this.m_ToolUtilityClass.SetOptionSetAttribute(ref aFeeToCreated, "new_pay_way", 100000004); // 100000004 = 未知、100000005=LinePay

                // 收費單付款狀態
                this.m_ToolUtilityClass.SetOptionSetAttribute(ref aFeeToCreated, "new_pay_status", 100000000); // 100000000 = 新建立

                // 收費單上課紀錄關聯
                this.m_ToolUtilityClass.SetEntityLookUpAttribute(ref aFeeToCreated, "new_stor_lessons_new_fee", "new_stor_lessons", NewStorLessonId);

                // 收費單收費日期
                this.m_ToolUtilityClass.SetEntityDateTimeAttribute(ref aFeeToCreated, "new_pay_date", DateTime.Now.ToLocalTime());

                // 永豐金流 QPay
                // 信用卡訂單編號
                this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeToCreated, "new_q_pay_card_order_no", CreatedCardOrder.OrderNo);
                // 虛擬帳號訂單編號
                this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeToCreated, "new_q_pay_order_atm_no", CreatedAtmOrder.OrderNo);

                // 轉帳/匯款編號
                this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeToCreated, "new_atm_pay_number", CreatedAtmOrder.ATMParam.AtmPayNo);

                // Line Pay
                //this.m_ToolUtilityClass.SetEntityIntAttribute(ref aFeeToCreated, "new_transaction_id", (int)aReserveResponse.Info.TransactionId);

                //// 交易識別碼
                //this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeToCreated, "new_transaction_id_string", aReserveResponse.Info.TransactionId.ToString());
                ////付款憑證
                //this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeToCreated, "new_payment_access_token", aReserveResponse.Info.PaymentAccessToken);
                //// 付款網頁
                //this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeToCreated, "new_payment_url_web", aReserveResponse.Info.PaymentUrl.Web);
                //// 付款應用
                //this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeToCreated, "new_payment_url_app", aReserveResponse.Info.PaymentUrl.App);

                #endregion
            }
            catch (System.Exception e)
            {
                String ErrorString = "ERROR : FullName = " + this.GetType().FullName.ToString() + " , Time = " + DateTime.Now.ToString() + " , Description = " + e.ToString();

                //Monitor.Exit(this);
                throw e;
            }
        }
        #endregion
        #region 永豐金流工具區
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

            QryOrderPay retObj = QPayToolkit.OrderPayQuery(orderPayQueryReq);

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
    }

}
