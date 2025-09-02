using ChurchReport.Models;
using Line.Messaging;
using Microsoft.Extensions.Configuration;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk;
using QPay.Domain;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using ToolUtilityNameSpace;
using UserProfile = Line.Messaging.UserProfile;

namespace ChurchReport.WebServiceConnector
{
    public class MyPayProcessor
    {
        #region 資料區
        private ToolUtilityClass m_ToolUtilityClass { get; set; }
        private IConfiguration m_Configuration { get; set; }

        #region LINE Bot 設定
        // 聖谷行道會
        private const String CHANNEL_ACCESS_TOKEN = @"OMjL23DpFRDgphgN7JdzA7uCpv1wb4hXtsGh4FzxP8tHzeMyYOr/ry3BBqaRNJpVUhR6wPHLN4Wa4QiG5i3P5T/Y07swP5OjfCz9DKwTYC7T4mPb8x54pwtcqK1lIdgNm6skdZnu99fBsupEcbZLBAdB04t89/1O/w1cDnyilFU=";

        //private LinePayClient m_LinePayClient { get; }

        private LineMessagingClient m_LineMessagingClient { get; }
        private PushUtility m_PushUtility { get; }
        private ReplyUtility m_ReplyUtility { get; }

        //private LineNotifyUtility m_LineNotifyUtility = new LineNotifyUtility();
        #endregion
        #endregion
        #region 初始化
        public MyPayProcessor()
        {
            m_ToolUtilityClass = new ToolUtilityClass("DYNAMICS365", "ymllc");
            // 讀取 appsettings.json 配置
            IConfigurationBuilder builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .AddEnvironmentVariables();

            m_Configuration = builder.Build();
        }
        #endregion
        #region 高鉅金流 PayPage 回傳處理
        /// <summary>
        /// 驗證高鉅金流回傳的 Hash 簽名
        /// </summary>
        /// <param name="returnModel">回傳資料</param>
        /// <returns>驗證結果</returns>
        public bool VerifyMyPayHash(MyPayReturnModel returnModel)
        {
            try
            {
                m_ToolUtilityClass = new ToolUtilityClass("DYNAMICS365", returnModel.echo_1);

                string key = m_Configuration["MyPay:Key"];
                string iv = m_Configuration["MyPay:IV"];

                if (string.IsNullOrEmpty(key) || string.IsNullOrEmpty(iv))
                {
                    String ErrorString = $"ERROR: MyPay Key 或 IV 設定為空 - {DateTime.Now}";
                    return false;
                }

                // 根據高鉅金流文檔的簽名計算規則
                // 簽名組合：KEY + transaction_id + order_id + state + IV
                string rawData = $"{key}{returnModel.transaction_id}{returnModel.order_id}{returnModel.state}{iv}";

                // 使用 SHA256 計算 Hash
                using (SHA256 sha256 = SHA256.Create())
                {
                    byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(rawData));
                    StringBuilder hashBuilder = new StringBuilder();

                    foreach (byte b in bytes)
                    {
                        hashBuilder.Append(b.ToString("x2"));
                    }

                    string calculatedHash = hashBuilder.ToString().ToUpper();
                    return calculatedHash.Equals(returnModel.hash, StringComparison.OrdinalIgnoreCase);
                }
            }
            catch (Exception ex)
            {
                String ErrorString = $"ERROR: VerifyMyPayHash - {DateTime.Now} - {ex}";
                return false;
            }
        }

        /// <summary>
        /// 處理高鉅金流回傳資訊並更新 Dynamics 365
        /// </summary>
        /// <param name="returnModel">回傳資料</param>
        /// <returns>處理結果</returns>
        public async Task<bool> ProcessMyPayReturn(MyPayReturnModel returnModel)
        {
            try
            {
                // 嘗試解析 order_id 成 Guid
                if (!Guid.TryParse(returnModel.order_id, out Guid entityId))
                {
                    String ErrorString = $"ERROR: 無法解析 order_id 為 Guid: {returnModel.order_id}";
                    return false;
                }

                // 先查詢收費單
                Entity entity = this.m_ToolUtilityClass.RetrieveEntity("new_fee", entityId);
                string entityType = "new_fee";

                // 如果找不到收費單，嘗試查詢認獻單
                if (entity == null)
                {
                    entity = this.m_ToolUtilityClass.RetrieveEntity("new_dedication_booking", entityId);
                    entityType = "new_dedication_booking";

                    if (entity == null)
                    {
                        String ErrorString = $"ERROR: 找不到對應的收費單或認獻單: {returnModel.order_id}";
                        return false;
                    }
                }

                // 檢查是否已處理過此交易 (冪等性處理)
                string existingTransactionId = this.m_ToolUtilityClass.GetEntityStringAttribute(entity, "new_mypay_transaction_id");
                if (!string.IsNullOrEmpty(existingTransactionId) && existingTransactionId == returnModel.transaction_id)
                {
                    // 已處理過此交易，直接回傳成功
                    return true;
                }

                // 記錄交易ID (用於避免重複處理)
                this.m_ToolUtilityClass.SetEntityStringAttribute(ref entity, "new_mypay_transaction_id", returnModel.transaction_id);

                // 根據交易結果進行不同處理
                if (returnModel.state == "1") // 交易成功
                {
                    await ProcessSuccessfulMyPayReturn(entity, entityType, returnModel);
                }
                else // 交易失敗
                {
                    await ProcessFailedMyPayReturn(entity, entityType, returnModel);
                }

                // 更新實體到 Dynamics 365
                this.m_ToolUtilityClass.UpdateEntity(entity);

                // 發送LINE通知 (可選)
                await SendMyPaymentNotification(entity, entityType, returnModel);

                return true;
            }
            catch (Exception ex)
            {
                String ErrorString = $"ERROR: ProcessMyPayReturn - {DateTime.Now} - {ex}";
                return false;
            }
        }

        /// <summary>
        /// 處理高鉅金流付款成功的情況
        /// </summary>
        /// <param name="entity">要更新的實體</param>
        /// <param name="entityType">實體類型</param>
        /// <param name="returnModel">回傳資料</param>
        private async Task ProcessSuccessfulMyPayReturn(Entity entity, string entityType, MyPayReturnModel returnModel)
        {
            if (entityType == "new_fee")
            {
                // 處理收費單付款成功
                // 更新付款狀態為已付款
                SetPayStatus("信用卡已繳費", ref entity);

                // 更新實收金額
                if (!string.IsNullOrEmpty(returnModel.cost) && decimal.TryParse(returnModel.cost, out decimal costValue) && costValue > 0)
                {
                    this.m_ToolUtilityClass.SetEntityMoneyAttribute(ref entity, "new_fee_really_paid", new Money(costValue));
                }
                else
                {
                    // 如果未回傳金額，使用應收金額
                    Money shouldPay = this.m_ToolUtilityClass.GetEntityMoneyAttribute(entity, "new_fee_shoud_pay");
                    if (shouldPay != null && shouldPay.Value > 0)
                    {
                        this.m_ToolUtilityClass.SetEntityMoneyAttribute(ref entity, "new_fee_really_paid", shouldPay);
                    }
                }

                // 更新付款日期
                this.m_ToolUtilityClass.SetEntityDateTimeAttribute(ref entity, "new_pay_date", DateTime.Now.ToLocalTime());
            }
            else if (entityType == "new_dedication_booking")
            {
                // 處理認獻單付款成功
                // 認獻單狀態設為已啟動
                this.m_ToolUtilityClass.SetOptionSetAttribute(ref entity, "new_dedication_booking_status", 100000001); // 已啟動
            }

            // 更新備註，記錄成功訊息
            string currentNote = this.m_ToolUtilityClass.GetEntityStringAttribute(entity, "new_explain") ?? "";
            string newNote = $"{currentNote}\n[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 高鉅金流付款成功\n" +
                 $"交易號: {returnModel.transaction_id}\n" +
                 $"金額: {(string.IsNullOrEmpty(returnModel.cost) ? "0" : returnModel.cost)} 元\n" +
                 $"訊息: {returnModel.msg}";
            this.m_ToolUtilityClass.SetEntityStringAttribute(ref entity, "new_explain", newNote);
        }

        /// <summary>
        /// 處理高鉅金流付款失敗的情況
        /// </summary>
        /// <param name="entity">要更新的實體</param>
        /// <param name="entityType">實體類型</param>
        /// <param name="returnModel">回傳資料</param>
        private async Task ProcessFailedMyPayReturn(Entity entity, string entityType, MyPayReturnModel returnModel)
        {
            // 更新備註，記錄失敗原因
            string currentNote = this.m_ToolUtilityClass.GetEntityStringAttribute(entity, "new_explain") ?? "";
            string newNote = $"{currentNote}\n[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 高鉅金流付款失敗\n" +
                           $"交易號: {returnModel.transaction_id}\n" +
                           $"失敗原因: {returnModel.msg}";
            this.m_ToolUtilityClass.SetEntityStringAttribute(ref entity, "new_explain", newNote);

            // 可以選擇是否要將付款狀態設為失敗，或保持原狀態
            // 如果有付款失敗的狀態選項，可以在這裡設定
        }

        /// <summary>
        /// 發送高鉅金流付款結果通知
        /// </summary>
        /// <param name="entity">實體</param>
        /// <param name="entityType">實體類型</param>
        /// <param name="returnModel">回傳資料</param>
        private async Task SendMyPaymentNotification(Entity entity, string entityType, MyPayReturnModel returnModel)
        {
            try
            {
                // 取得關聯聯絡人
                Guid contactId = Guid.Empty;

                if (entityType == "new_fee")
                {
                    contactId = this.m_ToolUtilityClass.GetEntityLookupAttribute(entity, "new_contact_new_fee");
                }
                else if (entityType == "new_dedication_booking")
                {
                    contactId = this.m_ToolUtilityClass.GetEntityLookupAttribute(entity, "new_contact_new_dedication_booking");
                }

                if (contactId != Guid.Empty)
                {
                    Entity contact = this.m_ToolUtilityClass.RetrieveEntity("contact", contactId);
                    if (contact != null)
                    {
                        string lineId = this.m_ToolUtilityClass.GetEntityStringAttribute(contact, "new_lineid");

                        if (!string.IsNullOrEmpty(lineId))
                        {
                            string message;
                            if (returnModel.state == "1")
                            {
                                // 付款成功訊息
                                message = $"您好，您的奉獻已經成功完成！\n" +
                                          $"交易號: {returnModel.transaction_id}\n" +
                                          $"金額: {(string.IsNullOrEmpty(returnModel.cost) ? "0" : returnModel.cost)} 元\n" +
                                          $"感謝您的奉獻！";
                            }
                            else
                            {
                                // 付款失敗訊息
                                message = $"您好，您的奉獻交易處理失敗。\n" +
                                         $"原因: {returnModel.msg}\n" +
                                         $"請稍後再試或聯繫教會辦公室。";
                            }

                            await m_PushUtility.SendMessage(lineId, message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 推送失敗不影響主流程，只記錄錯誤
                String ErrorString = $"ERROR: SendMyPaymentNotification - {DateTime.Now} - {ex}";
            }
        }
        #region 發展中
        //private async Task ProcessSuccessfulPayment(Entity entity, MyPayReturnModel returnModel)
        //{
        //    // 處理基本交易資訊
        //    ProcessBasicTransactionInfo(entity, returnModel);

        //    // 處理信用卡資訊
        //    ProcessCreditCardInfo(entity, returnModel);

        //    // 處理虛擬帳號資訊
        //    ProcessVirtualAccountInfo(entity, returnModel);

        //    // 處理定期定額資訊
        //    ProcessRecurringPaymentInfo(entity, returnModel);

        //    // 處理發票資訊
        //    ProcessInvoiceInfo(entity, returnModel);
        //}

        //private void ProcessBasicTransactionInfo(Entity entity, MyPayReturnModel returnModel)
        //{
        //    // 更新付款狀態
        //    if (!string.IsNullOrEmpty(returnModel.pfn))
        //    {
        //        SetPaymentMethod(returnModel.pfn, ref entity);
        //    }

        //    // 更新交易金額
        //    if (!string.IsNullOrEmpty(returnModel.cost))
        //    {
        //        if (int.TryParse(returnModel.cost, out int amount))
        //        {
        //            this.m_ToolUtilityClass.SetEntityMoneyAttribute(ref entity, "new_fee_really_paid", new Money(amount));
        //        }
        //    }

        //    // 更新交易完成時間
        //    if (!string.IsNullOrEmpty(returnModel.finishtime))
        //    {
        //        if (DateTime.TryParseExact(returnModel.finishtime, "yyyyMMddHHmmss", null, DateTimeStyles.None, out DateTime finishTime))
        //        {
        //            this.m_ToolUtilityClass.SetEntityDateTimeAttribute(ref entity, "new_pay_date", finishTime);
        //        }
        //    }

        //    // 記錄交易資訊
        //    string transactionInfo = $"高鉅金流交易資訊:\n" +
        //                           $"交易流水號: {returnModel.uid}\n" +
        //                           $"交易完成時間: {returnModel.finishtime}\n" +
        //                           $"金融服務商: {returnModel.supplier_name}\n" +
        //                           $"實際金額: {returnModel.actual_cost} {returnModel.actual_currency}";

        //    this.m_ToolUtilityClass.SetEntityStringAttribute(ref entity, "new_mypay_transaction_details", transactionInfo);
        //}

        //private void ProcessCreditCardInfo(Entity entity, MyPayReturnModel returnModel)
        //{
        //    if (!string.IsNullOrEmpty(returnModel.card_type))
        //    {
        //        // 儲存信用卡資訊
        //        string cardInfo = $"卡別: {returnModel.card_type}\n" +
        //                         $"發卡行: {returnModel.issuing_bank}\n" +
        //                         $"授權碼: {returnModel.acode}\n" +
        //                         $"卡號: ****-****-****-{returnModel.cardno}";

        //        this.m_ToolUtilityClass.SetEntityStringAttribute(ref entity, "new_mypay_card_info", cardInfo);

        //        // 處理分期資訊
        //        if (!string.IsNullOrEmpty(returnModel.installment))
        //        {
        //            this.m_ToolUtilityClass.SetEntityStringAttribute(ref entity, "new_installment_info", returnModel.installment);
        //        }

        //        // 處理紅利資訊
        //        if (!string.IsNullOrEmpty(returnModel.redeem))
        //        {
        //            this.m_ToolUtilityClass.SetEntityStringAttribute(ref entity, "new_redeem_info", returnModel.redeem);
        //        }
        //    }
        //}

        //private void ProcessVirtualAccountInfo(Entity entity, MyPayReturnModel returnModel)
        //{
        //    if (!string.IsNullOrEmpty(returnModel.bank_id))
        //    {
        //        // 儲存虛擬帳號資訊
        //        string atmInfo = $"銀行代碼: {returnModel.bank_id}\n" +
        //                        $"到期日: {returnModel.expired_date}\n" +
        //                        $"帳號資訊: {returnModel.result_content}";

        //        this.m_ToolUtilityClass.SetEntityStringAttribute(ref entity, "new_virtual_account_info", atmInfo);
        //    }
        //}

        //private void ProcessRecurringPaymentInfo(Entity entity, MyPayReturnModel returnModel)
        //{
        //    if (!string.IsNullOrEmpty(returnModel.group_id))
        //    {
        //        // 儲存定期定額資訊
        //        this.m_ToolUtilityClass.SetEntityStringAttribute(ref entity, "new_recurring_group_id", returnModel.group_id);
        //        this.m_ToolUtilityClass.SetEntityStringAttribute(ref entity, "new_recurring_payment_name", returnModel.payment_name);
        //        this.m_ToolUtilityClass.SetEntityStringAttribute(ref entity, "new_recurring_periods", returnModel.nois);
        //    }
        //}

        //private void ProcessInvoiceInfo(Entity entity, MyPayReturnModel returnModel)
        //{
        //    // 處理發票相關資訊 (根據您的 MyPayReturnModel 中的發票欄位)
        //    if (!string.IsNullOrEmpty(returnModel.invoice_number))
        //    {
        //        this.m_ToolUtilityClass.SetEntityStringAttribute(ref entity, "new_invoice_number", returnModel.invoice_number);
        //    }
        //}
        #endregion
        #endregion
        #region 工具區
        public void SetPayMethod(String Value, ref Entity aFeeEntity)
        {
            // 收費單付款狀態
            switch (Value)
            {
                case "未知":
                    this.m_ToolUtilityClass.SetOptionSetAttribute(aFeeEntity, "new_pay_way", 100000004);
                    break;
                case "現金":
                    this.m_ToolUtilityClass.SetOptionSetAttribute(aFeeEntity, "new_pay_way", 100000000);
                    break;
                case "信用卡":
                    this.m_ToolUtilityClass.SetOptionSetAttribute(aFeeEntity, "new_pay_way", 100000001);
                    break;
                case "ATM轉帳/匯款":
                    this.m_ToolUtilityClass.SetOptionSetAttribute(aFeeEntity, "new_pay_way", 100000002);
                    break;
                case "銀行轉帳":
                    this.m_ToolUtilityClass.SetOptionSetAttribute(aFeeEntity, "new_pay_way", 100000006);
                    break;
                case "超商付款":
                    this.m_ToolUtilityClass.SetOptionSetAttribute(aFeeEntity, "new_pay_way", 100000004);
                    break;
                case "行動支付":
                    this.m_ToolUtilityClass.SetOptionSetAttribute(aFeeEntity, "new_pay_way", 100000007);
                    break;
                case "銀聯卡":
                    this.m_ToolUtilityClass.SetOptionSetAttribute(aFeeEntity, "new_pay_way", 100000008);
                    break;
                case "LinePay":
                    this.m_ToolUtilityClass.SetOptionSetAttribute(aFeeEntity, "new_pay_way", 100000005);
                    break;
                default:
                    this.m_ToolUtilityClass.SetOptionSetAttribute(aFeeEntity, "new_pay_way", 100000004);
                    break;

            }
        }
        public void SetPayStatus(String Value, ref Entity aFeeEntity)
        {

            switch (Value)
            {
                case "新建立":
                    this.m_ToolUtilityClass.SetOptionSetAttribute(aFeeEntity, "new_pay_status", 100000000);
                    break;
                case "信用卡已繳費":
                    this.m_ToolUtilityClass.SetOptionSetAttribute(aFeeEntity, "new_pay_status", 100000001);
                    break;
                case "ATM轉帳/匯款已繳費":
                    this.m_ToolUtilityClass.SetOptionSetAttribute(aFeeEntity, "new_pay_status", 100000002);
                    break;
                case "現金已繳費":
                    this.m_ToolUtilityClass.SetOptionSetAttribute(aFeeEntity, "new_pay_status", 100000003);
                    break;
                case "銀行轉帳已繳費":
                    this.m_ToolUtilityClass.SetOptionSetAttribute(aFeeEntity, "new_pay_status", 100000004);
                    break;
                default:
                    this.m_ToolUtilityClass.SetOptionSetAttribute(aFeeEntity, "new_pay_status", 100000000);
                    break;

            }
        }
        //public void SendGratitudeLineMessage(Entity aContact, QpayModel QpayModel)
        //{
        //    try
        //    {
        //        #region 非同步建立收費單
        //        String LineId = this.m_ToolUtilityClass.GetEntityStringAttribute(ref aContact, "new_lineid");

        //        if (LineId != "")
        //        {
        //            String GratitudeMessage =
        //                "敬收 " + m_ToolUtilityClass.GetEntityStringAttribute(ref aContact, "fullname") + " 奉獻" + Environment.NewLine +
        //                "日期 : " + QpayModel.DedicationDate.ToShortDateString() + Environment.NewLine +
        //                "類別 : " + QpayModel.Category + "  " + QpayModel.Others + Environment.NewLine +
        //                "付款方式: " + QpayModel.PayWay + Environment.NewLine +
        //                "金額 : " + QpayModel.Amount;

        //            m_PushUtility.SendMessage(LineId, GratitudeMessage);
        //        }

        //        #endregion
        //    }
        //    catch (System.Exception e)
        //    {
        //        String ErrorString = "ERROR : FullName = " + this.GetType().FullName.ToString() + " , Time = " + DateTime.Now.ToString() + " , Description = " + e.ToString();

        //        //Monitor.Exit(this);
        //        throw e;
        //    }
        //}

        /// <summary>
        /// 實現阿拉伯數字到大寫中文的轉換，金額轉為大寫金額
        /// </summary>
        /// <param name="LowerMoney"></param>
        /// <returns></returns>
        public string MoneyToChinese(string LowerMoney)

        {

            string functionReturnValue = null;

            bool IsNegative = false; // 是否是負數

            if (LowerMoney.Trim().Substring(0, 1) == "-")

            {

                // 是負數則先轉為正數

                LowerMoney = LowerMoney.Trim().Remove(0, 1);

                IsNegative = true;

            }

            string strLower = null;

            string strUpart = null;

            string strUpper = null;

            int iTemp = 0;

            // 保留兩位小數 123.489→123.49　　123.4→123.4

            LowerMoney = Math.Round(double.Parse(LowerMoney), 2).ToString();

            if (LowerMoney.IndexOf(".") > 0)

            {

                if (LowerMoney.IndexOf(".") == LowerMoney.Length - 2)

                {

                    LowerMoney = LowerMoney + "0";

                }

            }

            else

            {

                LowerMoney = LowerMoney + ".00";

            }

            strLower = LowerMoney;

            iTemp = 1;

            strUpper = "";

            while (iTemp <= strLower.Length)

            {

                switch (strLower.Substring(strLower.Length - iTemp, 1))

                {

                    case ".":

                        strUpart = "圓";

                        break;

                    case "0":

                        strUpart = "零";

                        break;

                    case "1":

                        strUpart = "壹";

                        break;

                    case "2":

                        strUpart = "貳";

                        break;

                    case "3":

                        strUpart = "叄";

                        break;

                    case "4":

                        strUpart = "肆";

                        break;

                    case "5":

                        strUpart = "伍";

                        break;

                    case "6":

                        strUpart = "陸";

                        break;

                    case "7":

                        strUpart = "柒";

                        break;

                    case "8":

                        strUpart = "捌";

                        break;

                    case "9":

                        strUpart = "玖";

                        break;

                }

                switch (iTemp)

                {

                    case 1:

                        strUpart = strUpart + "分";

                        break;

                    case 2:

                        strUpart = strUpart + "角";

                        break;

                    case 3:

                        strUpart = strUpart + "";

                        break;

                    case 4:

                        strUpart = strUpart + "";

                        break;

                    case 5:

                        strUpart = strUpart + "拾";

                        break;

                    case 6:

                        strUpart = strUpart + "佰";

                        break;

                    case 7:

                        strUpart = strUpart + "仟";

                        break;

                    case 8:

                        strUpart = strUpart + "萬";

                        break;

                    case 9:

                        strUpart = strUpart + "拾";

                        break;

                    case 10:

                        strUpart = strUpart + "佰";

                        break;

                    case 11:

                        strUpart = strUpart + "仟";

                        break;

                    case 12:

                        strUpart = strUpart + "億";

                        break;

                    case 13:

                        strUpart = strUpart + "拾";

                        break;

                    case 14:

                        strUpart = strUpart + "佰";

                        break;

                    case 15:

                        strUpart = strUpart + "仟";

                        break;

                    case 16:

                        strUpart = strUpart + "萬";

                        break;

                    default:

                        strUpart = strUpart + "";

                        break;

                }

                strUpper = strUpart + strUpper;

                iTemp = iTemp + 1;

            }

            strUpper = strUpper.Replace("零拾", "零");

            strUpper = strUpper.Replace("零佰", "零");

            strUpper = strUpper.Replace("零仟", "零");

            strUpper = strUpper.Replace("零零零", "零");

            strUpper = strUpper.Replace("零零", "零");

            strUpper = strUpper.Replace("零角零分", "整");

            strUpper = strUpper.Replace("零分", "整");

            strUpper = strUpper.Replace("零角", "零");

            strUpper = strUpper.Replace("零億零萬零圓", "億圓");

            strUpper = strUpper.Replace("億零萬零圓", "億圓");

            strUpper = strUpper.Replace("零億零萬", "億");

            strUpper = strUpper.Replace("零萬零圓", "萬圓");

            strUpper = strUpper.Replace("零億", "億");

            strUpper = strUpper.Replace("零萬", "萬");

            strUpper = strUpper.Replace("零圓", "圓");

            strUpper = strUpper.Replace("零零", "零");

            // 對壹圓以下的金額的處理

            if (strUpper.Substring(0, 1) == "圓")

            {

                strUpper = strUpper.Substring(1, strUpper.Length - 1);

            }

            if (strUpper.Substring(0, 1) == "零")

            {

                strUpper = strUpper.Substring(1, strUpper.Length - 1);

            }

            if (strUpper.Substring(0, 1) == "角")

            {

                strUpper = strUpper.Substring(1, strUpper.Length - 1);

            }

            if (strUpper.Substring(0, 1) == "分")

            {

                strUpper = strUpper.Substring(1, strUpper.Length - 1);

            }

            if (strUpper.Substring(0, 1) == "整")

            {

                strUpper = "零圓整";

            }

            functionReturnValue = strUpper;

            if (IsNegative == true)

            {

                return "負" + functionReturnValue;

            }

            else

            {

                return functionReturnValue;

            }

        }

        private int TransferToDeductTotalNum(string DeductTotalNumber)
        {
            switch (DeductTotalNumber)
            {
                case "3個月":
                    return 3;
                case "6個月":
                    return 6;
                case "12個月":
                    return 12;
                case "18個月":
                    return 18;
                case "24個月":
                    return 24;
                default:
                    return 0;
            }
        }

        #endregion
    }
}
