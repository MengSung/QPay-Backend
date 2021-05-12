using Line.Messaging;
using QPayBackend.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Xrm.Sdk;
using QPay.Domain;
using System;
using System.Collections.Generic;
using ToolUtilityNameSpace;

namespace QPayBackend.Tools
{
    public class QPayFeeProcessor : Controller, IDisposable
    {
        #region 資料區
        private LineMessagingClient m_LineMessagingClient { get; set; }

        private PushUtility m_PushUtility { get; set; }

        private ToolUtilityClass m_ToolUtilityClass { get; set; }

        // 胡夢嵩回傳　EXCEPTION　專用的ＩＤ
        private const String MENGSUNG_LINE_ID = @"U7638e4ed509708a3573ba6d69970583d";

        // 音訊教會-雲端除錯用
        private const String CHANNEL_ACCESS_TOKEN = @"g1jtWWNkjbH3OCh1cKoRvPBUkCJIygNuvV/neHXR9I4J5GBgVE85inaIaTcT4AAZ1qCuqrqJXDawrUweyBqLcX97GGokXnTRQ6MxjXAutd5Yr2FkPsZnq6kMelc/C+mqNUHaVUKFAuvTD8JvXbNmpAdB04t89/1O/w1cDnyilFU=";
        #endregion
        #region 初始化
        public QPayFeeProcessor()
        {
        }
        #endregion
        #region 釋放記憶體
        private bool _disposed = false;

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                m_ToolUtilityClass.Dispose();
            }

            _disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~QPayFeeProcessor()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of
            // readability and maintainability.
            Dispose(false);
        }
        #endregion
        #region 主程式
        public JsonResult QPayBackendUrl(QryOrderPay aQryOrderPay)
        {
            try
            {
                if (aQryOrderPay.TSResultContent.Param2 != "")
                {
                    m_ToolUtilityClass = new ToolUtilityClass("DYNAMICS365", aQryOrderPay.TSResultContent.Param2);
                    this.m_LineMessagingClient = new LineMessagingClient(ConvertOrganzitionToChannelAccessToken(aQryOrderPay.TSResultContent.Param2));
                    //this.m_LineMessagingClient = new LineMessagingClient(CHANNEL_ACCESS_TOKEN);
                    m_PushUtility = new PushUtility(m_LineMessagingClient);
                }
                else
                {
                    return Json(new Dictionary<string, string>() { { "Status", "S" } });
                }

                #region// 取得收費單
                //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "003");
                Entity aFeeEntity = this.m_ToolUtilityClass.RetrieveEntity("new_fee", new Guid(aQryOrderPay.TSResultContent.Param1));
                //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "004");
                #endregion
                if (aFeeEntity != null)
                {
                    // 有找到收費單
                    // 收費單付款狀態
                    // 新建立 100000000
                    // 信用卡已繳費 100000001
                    // ATM轉帳/匯款已繳費 100000002
                    // 現金已繳費 100000003

                    // 取得付款人
                    //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "005");

                    Entity aContact = this.m_ToolUtilityClass.RetrieveEntity("contact", this.m_ToolUtilityClass.GetEntityLookupAttribute(aFeeEntity, "new_contact_new_fee"));
                    // 取得付款人姓名
                    String aFullName = this.m_ToolUtilityClass.GetEntityStringAttribute(aContact, "fullname");
                    // 取得付款人 Line Id
                    String UserLineId = this.m_ToolUtilityClass.GetEntityStringAttribute(aContact, "new_lineid");


                    #region// 收費單描述說明
                    String Description = "";
                    if (aQryOrderPay.TSResultContent.OrderNo.StartsWith("C"))
                    {
                        #region 信用卡
                        Description =
                               "姓名     : " + aFullName + Environment.NewLine +
                               "日期     : " + DateTime.Now.ToLocalTime().ToString() + Environment.NewLine +
                               "訂單編號 : " + aQryOrderPay.TSResultContent.OrderNo + Environment.NewLine +
                               "實收金額 : " + ((int)Convert.ToUInt32(aQryOrderPay.TSResultContent.Amount) / 100).ToString() + "元" + Environment.NewLine +
                               "付款方式 : " + "信用卡" + Environment.NewLine +
                               "程式呼叫 : " + aQryOrderPay.Description + Environment.NewLine +
                               "交易結果 : " + aQryOrderPay.TSResultContent.Description + Environment.NewLine +
                               "--------------------" + Environment.NewLine;
                        #endregion
                    }
                    else
                    {
                        #region ATM轉帳/匯款
                        Description =
                               "姓名     : " + aFullName + Environment.NewLine +
                               "訂單編號 : " + aQryOrderPay.TSResultContent.OrderNo + Environment.NewLine +
                               "日期     : " + DateTime.Now.ToLocalTime().ToString() + Environment.NewLine +
                               "實收金額 : " + ((int)Convert.ToUInt32(aQryOrderPay.TSResultContent.Amount) / 100).ToString() + "元" + Environment.NewLine +
                               "付款方式 : " + "ATM轉帳/匯款" + Environment.NewLine +
                               "程式呼叫 : " + aQryOrderPay.Description + Environment.NewLine +
                               "交易結果 : " + aQryOrderPay.TSResultContent.Description + Environment.NewLine +
                               "--------------------" + Environment.NewLine;
                        #endregion
                    }
                    #endregion
                    if (aQryOrderPay.Status == "S" && aQryOrderPay.TSResultContent.Status == "S")
                    {
                        if (aQryOrderPay.TSResultContent.OrderNo.StartsWith("C"))
                        {
                            // 處理信用卡
                            ////this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "007");

                            if (this.m_ToolUtilityClass.GetEntityStringAttribute(aFeeEntity, "new_payment_records").Contains(aQryOrderPay.TSResultContent.OrderNo) != true && this.m_ToolUtilityClass.GetOptionSetAttribute(ref aFeeEntity, "new_pay_status") == 100000000)
                            {
                                #region// 付款狀態 等於 "新建立"
                                #region 信用卡會回傳2次，一次是RETURN_URL、一次是BACKEND_URL，為免收費單紀錄信用卡兩次，所以如果這裡沒有信用卡訂單編號或是等於 "新建立"，才要進行處理了
                                // 收費單付款日期
                                this.m_ToolUtilityClass.SetEntityDateTimeAttribute(ref aFeeEntity, "new_pay_date", DateTime.Now.ToLocalTime());
                                // 收費單總共實收金額
                                Money aTotalPaid = new Money(Convert.ToUInt32(this.m_ToolUtilityClass.GetEntityMoneyAttribute(ref aFeeEntity, "new_fee_really_paid").Value + new Money((int)Convert.ToUInt32(aQryOrderPay.TSResultContent.Amount) / 100).Value));
                                this.m_ToolUtilityClass.SetEntityMoneyAttribute(ref aFeeEntity, "new_fee_really_paid", aTotalPaid);
                                // 收費單實現阿拉伯數字到大寫中文的轉換，金額轉為大寫金額
                                this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeEntity, "new_big_chinese_number", MoneyToChinese((Convert.ToUInt32(aQryOrderPay.TSResultContent.Amount) / 100).ToString()));
                                // 收費單付款方式
                                this.m_ToolUtilityClass.SetOptionSetAttribute(ref aFeeEntity, "new_pay_way", 100000001); // 100000001 = 信用卡
                                // 收費單付款狀態
                                this.m_ToolUtilityClass.SetOptionSetAttribute(ref aFeeEntity, "new_pay_status", 100000001); // 100000001 = 信用卡已繳費
                                // 收費單說明
                                String aOriginalDescription = this.m_ToolUtilityClass.GetEntityStringAttribute(ref aFeeEntity, "new_description");
                                this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeEntity, "new_description", aOriginalDescription + Description);
                                // 付款紀錄
                                String aPaymentRecords =
                                        this.m_ToolUtilityClass.GetEntityStringAttribute(aFeeEntity, "new_payment_records") +
                                        DateTime.Now.ToString() +
                                        ": BackendUrl => 信用卡訂單編號= " + aQryOrderPay.TSResultContent.OrderNo +
                                        "，金額:" + ((int)Convert.ToUInt32(aQryOrderPay.TSResultContent.Amount) / 100).ToString() +
                                        Environment.NewLine;
                                this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeEntity, "new_payment_records", aPaymentRecords);

                                if (aQryOrderPay.TSResultContent.OrderNo.StartsWith("C"))
                                {
                                    // 已付款信用卡訂單編號
                                    this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeEntity, "new_q_paid_card_order_no", aQryOrderPay.TSResultContent.OrderNo);
                                }

                                // 更新收費單
                                this.m_ToolUtilityClass.UpdateEntity(ref aFeeEntity);

                                #region// 取得上課紀錄單，更新報名狀態
                                Guid aStorLessonsId = this.m_ToolUtilityClass.GetEntityLookupAttribute(ref aFeeEntity, "new_stor_lessons_new_fee");
                                if (aStorLessonsId != Guid.Empty)
                                {
                                    Entity aStorLessons = this.m_ToolUtilityClass.RetrieveEntity("new_stor_lessons", aStorLessonsId);

                                    #region 報名狀態
                                    if (aQryOrderPay.TSResultContent.Param2 == "yhchurch" || aQryOrderPay.TSResultContent.Param2 == "imchurch" || aQryOrderPay.TSResultContent.Param2 == "thevictory")
                                    {
                                        // 有小組長審核的教會=>報名成功:永和禮拜堂
                                        this.m_ToolUtilityClass.SetOptionSetAttribute(ref aStorLessons, "new_enroll_status", 100000008);
                                    }
                                    else
                                    {
                                        // 不需要小組長審核的教會
                                        this.m_ToolUtilityClass.SetOptionSetAttribute(ref aStorLessons, "new_enroll_status", 100000001);
                                    }
                                    #endregion

                                    this.m_ToolUtilityClass.UpdateEntity(ref aStorLessons);
                                }
                                #endregion

                                #region// 設定連絡人信用卡資訊
                                if (aQryOrderPay.TSResultContent.CCToken != "")
                                {
                                    String VisaInfo = this.m_ToolUtilityClass.GetEntityStringAttribute(ref aContact, "new_visa_info");

                                    if (IsCreditCardInList(aContact, aQryOrderPay) != true)
                                    {
                                        VisaInfo =
                                                aQryOrderPay.TSResultContent.CCToken + "，" +
                                                aQryOrderPay.TSResultContent.LeftCCNo + "，" +
                                                aQryOrderPay.TSResultContent.RightCCNo + "，" +
                                                //aQryOrderPay.TSResultContent.AuthCode + "，" +
                                                aQryOrderPay.TSResultContent.CCExpDate +
                                                "|" + VisaInfo;


                                        this.m_ToolUtilityClass.SetEntityStringAttribute(ref aContact, "new_visa_info", VisaInfo);

                                        this.m_ToolUtilityClass.UpdateEntity(ref aContact);
                                    }
                                }
                                #endregion

                                // LINE 通知付款人
                                this.m_PushUtility.SendMessage(UserLineId, "信用卡付款結果成功!" + Environment.NewLine + Description);
                                #endregion
                                #endregion
                            }
                        }
                        else
                        {
                            // 處理帳號訂單 : ATM轉帳/匯款
                            //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "008");
                            if (this.m_ToolUtilityClass.GetOptionSetAttribute(ref aFeeEntity, "new_pay_status") == 100000000)
                            {
                                //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "008.1");
                                #region// 付款狀態 等於 "新建立"

                                // 收費單付款日期
                                this.m_ToolUtilityClass.SetEntityDateTimeAttribute(ref aFeeEntity, "new_pay_date", DateTime.Now.ToLocalTime());

                                // 收費單總共實收金額
                                Money aTotalPaid = new Money(Convert.ToUInt32(this.m_ToolUtilityClass.GetEntityMoneyAttribute(ref aFeeEntity, "new_fee_really_paid").Value + new Money((int)Convert.ToUInt32(aQryOrderPay.TSResultContent.Amount) / 100).Value));
                                this.m_ToolUtilityClass.SetEntityMoneyAttribute(ref aFeeEntity, "new_fee_really_paid", aTotalPaid);
                                // 收費單實現阿拉伯數字到大寫中文的轉換，金額轉為大寫金額
                                this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeEntity, "new_big_chinese_number", MoneyToChinese((Convert.ToUInt32(aQryOrderPay.TSResultContent.Amount) / 100).ToString()));
                                // 收費單付款方式
                                this.m_ToolUtilityClass.SetOptionSetAttribute(ref aFeeEntity, "new_pay_way", 100000002); // 100000002 = ATM轉帳/匯款
                                // 收費單付款狀態
                                this.m_ToolUtilityClass.SetOptionSetAttribute(ref aFeeEntity, "new_pay_status", 100000002); // 100000002 = ATM轉帳/匯款已繳費
                                // 收費單說明
                                String aOriginalDescription = this.m_ToolUtilityClass.GetEntityStringAttribute(ref aFeeEntity, "new_description");
                                this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeEntity, "new_description", aOriginalDescription + Description);
                                //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "008.2");
                                // 付款紀錄
                                String aPaymentRecords =
                                        this.m_ToolUtilityClass.GetEntityStringAttribute(aFeeEntity, "new_payment_records") +
                                        DateTime.Now.ToString() +
                                        ": BackendUrl => ATM轉帳/匯款訂單編號= " + aQryOrderPay.TSResultContent.OrderNo +
                                        "，金額:" + ((int)Convert.ToUInt32(aQryOrderPay.TSResultContent.Amount) / 100).ToString() +
                                        Environment.NewLine;
                                this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeEntity, "new_payment_records", aPaymentRecords);

                                //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "008.3");

                                if (aQryOrderPay.TSResultContent.OrderNo.StartsWith("A"))
                                {
                                    // 已付款虛擬帳號訂單編號
                                    this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeEntity, "new_q_paid_order_atm_no", aQryOrderPay.TSResultContent.OrderNo);
                                }

                                // 更新收費單
                                this.m_ToolUtilityClass.UpdateEntity(ref aFeeEntity);

                                #region// 取得上課紀錄單，更新報名狀態
                                Guid aStorLessonsId = this.m_ToolUtilityClass.GetEntityLookupAttribute(ref aFeeEntity, "new_stor_lessons_new_fee");
                                if (aStorLessonsId != Guid.Empty)
                                {
                                    Entity aStorLessons = this.m_ToolUtilityClass.RetrieveEntity("new_stor_lessons", aStorLessonsId);

                                    #region 報名狀態
                                    if (aQryOrderPay.TSResultContent.Param2 == "yhchurch" || aQryOrderPay.TSResultContent.Param2 == "imchurch" || aQryOrderPay.TSResultContent.Param2 == "thevictory")
                                    {
                                        // 有小組長審核的教會=>報名成功:永和禮拜堂 、 iM行動教會
                                        this.m_ToolUtilityClass.SetOptionSetAttribute(ref aStorLessons, "new_enroll_status", 100000008);
                                    }
                                    else
                                    {
                                        // 不需要小組長審核的教會
                                        this.m_ToolUtilityClass.SetOptionSetAttribute(ref aStorLessons, "new_enroll_status", 100000001);
                                    }
                                    #endregion

                                    this.m_ToolUtilityClass.UpdateEntity(ref aStorLessons);
                                }
                                #endregion

                                // LINE 通知付款人
                                //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "008.4 =>" + Description);
                                this.m_PushUtility.SendMessage(UserLineId, "ATM轉帳/匯款付款結果成功!" + Environment.NewLine + Description);
                                #endregion
                            }
                        }
                    }
                    else
                    {
                        // 收費單說明
                        String aOriginalDescription = this.m_ToolUtilityClass.GetEntityStringAttribute(ref aFeeEntity, "new_description");
                        this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeEntity, "new_description", aOriginalDescription + "付款結果失敗!" + Environment.NewLine + Description);

                        // 更新收費單
                        this.m_ToolUtilityClass.UpdateEntity(ref aFeeEntity);

                        // LINE 通知付款人
                        this.m_PushUtility.SendMessage(UserLineId, "付款失敗!" + Environment.NewLine + Description);
                    }
                }

                ////this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "009");

                return Json(new Dictionary<string, string>() { { "Status", "S" } });
            }
            catch (System.Exception e)
            {
                String ErrorString = "ERROR : FullName = " + this.GetType().FullName.ToString() + " , Time = " + DateTime.Now.ToString() + " , Description = " + e.ToString();

                if (this.m_PushUtility != null)
                {
                    this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, ErrorString);
                }

                return Json(new Dictionary<string, string>() { { "Status", "S" } });
            }

        }
        #endregion
        #region 商店代號對應區
        private string ConvertShopNoToOrganization(String ShopNo)
        {
            switch (ShopNo)
            {
                case "NA0149_001":
                    //return "chunghsiaochurch";
                    //return "yhchurchback";
                    return "jesus";
                case "DA1626_001":
                    return "yhchurch";
                case "DA2424_001":
                    // iM行動教會
                    return "imchurch";
                default:
                    return null;
            }
        }
        private string ConvertShopNoToChannelAccessToken(String ShopNo)
        {
            switch (ShopNo)
            {
                case "NA0149_001":
                    // 永和禮拜堂(公司研發)
                    //return @"Z821JyND95uiABqED/bwOcTyCkHMcp92JBDYJn/oefwaIseWFyLSDKtTeB+SqMI1kquELAvJ7TSN+EDhl7WGgfFLgT9zehh8+3ocAQEKmfCzTzio5xoHKxfQzrvlXmCtp7wfm4vuPT33dr7tBJrkOAdB04t89/1O/w1cDnyilFU=";
                    // 永和禮拜堂(雲端機房)
                    //return @"HeuLkSEF5CX7hdZo4956IPpgJNdb8VqRZeL1Gu37kFFm+1F7DObAGjfeVYaggzwjZ5H4qraesvquODt7Y81jbtspNZkEq5n3oLDG+G32xQsRx1jCobkABL/Z7RKjkSACNT6h72bPQXsVn9aCuI5OogdB04t89/1O/w1cDnyilFU=";
                    // 音訊教會
                    return @"g1jtWWNkjbH3OCh1cKoRvPBUkCJIygNuvV/neHXR9I4J5GBgVE85inaIaTcT4AAZ1qCuqrqJXDawrUweyBqLcX97GGokXnTRQ6MxjXAutd5Yr2FkPsZnq6kMelc/C+mqNUHaVUKFAuvTD8JvXbNmpAdB04t89/1O/w1cDnyilFU=";
                // 忠孝路長老教會
                //return @"aKS4zYeq2ZpqlLd4gslkWAyYuiC+B2f1noatF1VylPvkR2+mrvJ7mwnIIXtn2Pi117NBmNTmRZL5DO5ZMYaGCj/v9+fB6Zn9sel42Jr55PlegJdrtoSvPgm4fBso1tY/7H65+cOFDQxjqhdOU69qQAdB04t89/1O/w1cDnyilFU=";
                case "DA1626_001":
                    // 永和禮拜堂(公司研發)
                    //return @"Z821JyND95uiABqED/bwOcTyCkHMcp92JBDYJn/oefwaIseWFyLSDKtTeB+SqMI1kquELAvJ7TSN+EDhl7WGgfFLgT9zehh8+3ocAQEKmfCzTzio5xoHKxfQzrvlXmCtp7wfm4vuPT33dr7tBJrkOAdB04t89/1O/w1cDnyilFU=";
                    // 永和禮拜堂(雲端機房)
                    return @"HeuLkSEF5CX7hdZo4956IPpgJNdb8VqRZeL1Gu37kFFm+1F7DObAGjfeVYaggzwjZ5H4qraesvquODt7Y81jbtspNZkEq5n3oLDG+G32xQsRx1jCobkABL/Z7RKjkSACNT6h72bPQXsVn9aCuI5OogdB04t89/1O/w1cDnyilFU=";
                case "DA2424_001":
                    // iM行動教會
                    return @"XwSRWX0RxTtTvY/N6QZQ9YElOMH3OAxBf/3DAmWoXbIK3ymBsXEaU54owfdbPTQiQJPd10cWjC+JIWX6EvOCTbBdHmmJNC6xOOaioB91gPJPyDpl0IHQOQAzLA9J21zZ83SgIF6JwJbxC/8tSXv6RgdB04t89/1O/w1cDnyilFU=";
                default:
                    //return @"aKS4zYeq2ZpqlLd4gslkWAyYuiC+B2f1noatF1VylPvkR2+mrvJ7mwnIIXtn2Pi117NBmNTmRZL5DO5ZMYaGCj/v9+fB6Zn9sel42Jr55PlegJdrtoSvPgm4fBso1tY/7H65+cOFDQxjqhdOU69qQAdB04t89/1O/w1cDnyilFU=";
                    return @"HeuLkSEF5CX7hdZo4956IPpgJNdb8VqRZeL1Gu37kFFm+1F7DObAGjfeVYaggzwjZ5H4qraesvquODt7Y81jbtspNZkEq5n3oLDG+G32xQsRx1jCobkABL/Z7RKjkSACNT6h72bPQXsVn9aCuI5OogdB04t89/1O/w1cDnyilFU=";
            }
        }
        private string ConvertOrganzitionToChannelAccessToken(String Organzition)
        {
            //客製化
            switch (Organzition)
            {
                case "yhchurchback":
                    // 永和禮拜堂(公司研發)
                    return @"Z821JyND95uiABqED/bwOcTyCkHMcp92JBDYJn/oefwaIseWFyLSDKtTeB+SqMI1kquELAvJ7TSN+EDhl7WGgfFLgT9zehh8+3ocAQEKmfCzTzio5xoHKxfQzrvlXmCtp7wfm4vuPT33dr7tBJrkOAdB04t89/1O/w1cDnyilFU=";
                case "yhchurch":
                    // 永和禮拜堂(雲端機房)
                    return @"HeuLkSEF5CX7hdZo4956IPpgJNdb8VqRZeL1Gu37kFFm+1F7DObAGjfeVYaggzwjZ5H4qraesvquODt7Y81jbtspNZkEq5n3oLDG+G32xQsRx1jCobkABL/Z7RKjkSACNT6h72bPQXsVn9aCuI5OogdB04t89/1O/w1cDnyilFU=";
                case "chunghsiaochurch":
                    // 忠孝路長老教會(公司研發)
                    return @"aKS4zYeq2ZpqlLd4gslkWAyYuiC+B2f1noatF1VylPvkR2+mrvJ7mwnIIXtn2Pi117NBmNTmRZL5DO5ZMYaGCj/v9+fB6Zn9sel42Jr55PlegJdrtoSvPgm4fBso1tY/7H65+cOFDQxjqhdOU69qQAdB04t89/1O/w1cDnyilFU=";
                case "chunghsiao":
                    // 忠孝路長老教會(雲端機房)
                    return @"aKS4zYeq2ZpqlLd4gslkWAyYuiC+B2f1noatF1VylPvkR2+mrvJ7mwnIIXtn2Pi117NBmNTmRZL5DO5ZMYaGCj/v9+fB6Zn9sel42Jr55PlegJdrtoSvPgm4fBso1tY/7H65+cOFDQxjqhdOU69qQAdB04t89/1O/w1cDnyilFU=";
                case "jesus":
                    // 音訊教會
                    return @"g1jtWWNkjbH3OCh1cKoRvPBUkCJIygNuvV/neHXR9I4J5GBgVE85inaIaTcT4AAZ1qCuqrqJXDawrUweyBqLcX97GGokXnTRQ6MxjXAutd5Yr2FkPsZnq6kMelc/C+mqNUHaVUKFAuvTD8JvXbNmpAdB04t89/1O/w1cDnyilFU=";
                case "imchurch":
                    // iM行動教會
                    return @"XwSRWX0RxTtTvY/N6QZQ9YElOMH3OAxBf/3DAmWoXbIK3ymBsXEaU54owfdbPTQiQJPd10cWjC+JIWX6EvOCTbBdHmmJNC6xOOaioB91gPJPyDpl0IHQOQAzLA9J21zZ83SgIF6JwJbxC/8tSXv6RgdB04t89/1O/w1cDnyilFU=";
                case "imchurchback":
                    // iM行動教會(公司研發)
                    return @"YJ1LKtDZyfHwfkbqeHAk+pxNJNZBpOvI446h3brWHDqquFc2ElUCYaseqiW+pAKhwJspguAgGbOlKDymSjSTMydJn7JeY6CRmeyC2Am7urM3CNVNq/2JzAuQ2Vqc7lhPWx8qX5YxS3ve4NjcDceymQdB04t89/1O/w1cDnyilFU=";
                case "thevictory":
                    // 得勝靈糧堂(雲端機房)
                    return @"dhWNUj4LOTQFl10j0nvn+7/O3ffZkqfBz5+H6WKGoktwTpu32T+rdJYUfDSvT8HRz+VNkRcbttdJ74d81MecfD/q8AuUK5fhi8/eL9xFnDZBCCqLGP6q9lcZjvleoUXxN/OVfd2kcU3C4jk7sUP8pwdB04t89/1O/w1cDnyilFU=";
                case "thevictoryback":
                    // 得勝靈糧堂(公司研發)
                    return @"dhWNUj4LOTQFl10j0nvn+7/O3ffZkqfBz5+H6WKGoktwTpu32T+rdJYUfDSvT8HRz+VNkRcbttdJ74d81MecfD/q8AuUK5fhi8/eL9xFnDZBCCqLGP6q9lcZjvleoUXxN/OVfd2kcU3C4jk7sUP8pwdB04t89/1O/w1cDnyilFU=";
                default:
                    return @"aKS4zYeq2ZpqlLd4gslkWAyYuiC+B2f1noatF1VylPvkR2+mrvJ7mwnIIXtn2Pi117NBmNTmRZL5DO5ZMYaGCj/v9+fB6Zn9sel42Jr55PlegJdrtoSvPgm4fBso1tY/7H65+cOFDQxjqhdOU69qQAdB04t89/1O/w1cDnyilFU=";
            }
        }
        #endregion
        #region 工具區
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
        public bool IsCreditCardInList(Entity aContact, QryOrderPay aQryOrderPay)
        {
            #region// 取得連絡人信用卡資訊

            String VisaInfo = this.m_ToolUtilityClass.GetEntityStringAttribute(ref aContact, "new_visa_info");

            if (VisaInfo != "")
            {
                // 有儲存的信用卡
                String[] VisaInfoSplit = VisaInfo.Split('|');

                if (VisaInfoSplit.Length > 0)
                {
                    // 檢驗一個一個的信用卡是否重覆
                    foreach (String CreditCard in VisaInfoSplit)
                    {
                        String[] VisaCCTokenSplit = CreditCard.Split('，');

                        if (VisaCCTokenSplit.Length == 4)
                        {
                            if
                                (
                                    VisaCCTokenSplit[1] == aQryOrderPay.TSResultContent.LeftCCNo &&
                                    VisaCCTokenSplit[2] == aQryOrderPay.TSResultContent.RightCCNo &&
                                    VisaCCTokenSplit[3] == aQryOrderPay.TSResultContent.CCExpDate
                                )
                            {
                                // 有一樣的信用卡
                                return true;
                            }
                        }
                    }
                }
            }
            else
            {
                // 還沒有儲存的信用卡
                return true;
            }

            // 每個儲存的信用卡與目前要儲存的信用卡都不一樣
            return false;

            #endregion
        }

        #endregion
    }
}
