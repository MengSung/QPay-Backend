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

    public class QPayAtmWebhook : Controller
    {
        private LineMessagingClient m_LineMessagingClient { get; set; }

        private PushUtility m_PushUtility { get; set; }

        private QPayProcessor m_QPayProcessor { get; set; }

        private ToolUtilityClass m_ToolUtilityClass { get; set; }

        // 胡夢嵩回傳　EXCEPTION　專用的ＩＤ
        private const String MENGSUNG_LINE_ID = @"U7638e4ed509708a3573ba6d69970583d";

        public QPayAtmWebhook()
        {
            m_QPayProcessor = new QPayProcessor();

        }

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

        ~QPayAtmWebhook()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of
            // readability and maintainability.
            Dispose(false);
        }
        #endregion


        public JsonResult QPayBackendUrl([FromBody] BackendPostData aBackendPostData)
        {
            try
            {
                if (aBackendPostData.ShopNo != null)
                {
                    this.m_LineMessagingClient = new LineMessagingClient(ConvertShopNoToChannelAccessToken(aBackendPostData.ShopNo));
                }
                else
                {
                    this.m_LineMessagingClient = new LineMessagingClient(ConvertShopNoToChannelAccessToken(""));

                    this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "ShopNo值為NULL");

                    return Json(new Dictionary<string, string>()
                        {
                            { "Status", "S" }
                        });

                }
                m_PushUtility = new PushUtility(m_LineMessagingClient);

                //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "呼叫雲端ATM轉帳");

                m_ToolUtilityClass = new ToolUtilityClass("DYNAMICS365", ConvertShopNoToOrganization(aBackendPostData.ShopNo));

                //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "建立m_ToolUtilityClass完成");

                QryOrderPay aQryOrderPay = new QryOrderPay();

                aQryOrderPay = m_QPayProcessor.OrderPayQuery(aBackendPostData.PayToken);

                #region// 取得收費單
                EntityCollection FeeEntityCollection = m_ToolUtilityClass.RetrieveFeeByFetchXmlOrderNumber(aQryOrderPay.TSResultContent.OrderNo);
                Entity aFeeEntity;
                if (FeeEntityCollection.Entities.Count == 1)
                {
                    //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "有找到收費單");

                    aFeeEntity = this.m_ToolUtilityClass.RetrieveEntity("new_fee", FeeEntityCollection.Entities[0].Id);
                    // 收費單付款狀態
                    if (this.m_ToolUtilityClass.GetOptionSetAttribute(ref aFeeEntity, "new_pay_status") == 100000001) // 100000001 = 信用卡已繳費
                    {
                        return Json(new Dictionary<string, string>()
                        {
                            { "Status", "S" }
                        });
                    }
                }
                else
                {
                    //this.m_PushUtility.SendMessage( MENGSUNG_LINE_ID, "呼叫ATM轉帳-沒找到收費單");

                    return Json(new Dictionary<string, string>()
                    {
                        { "Status", "S" }
                    });
                }
                #endregion
                // 取得付款人
                Entity aContact = this.m_ToolUtilityClass.RetrieveEntity("contact", this.m_ToolUtilityClass.GetEntityLookupAttribute(aFeeEntity, "new_contact_new_fee"));
                // 取得付款人姓名
                String aFullName = this.m_ToolUtilityClass.GetEntityStringAttribute(aContact, "fullname");
                // 取得付款人 Line Id
                String UserLineId = this.m_ToolUtilityClass.GetEntityStringAttribute(aContact, "new_lineid");

                #region// 收費單描述說明
                String FeeDescription = this.m_ToolUtilityClass.GetEntityStringAttribute(ref aFeeEntity, "new_description");
                String Description = "";
                if (aQryOrderPay.TSResultContent.OrderNo.StartsWith("C"))
                {
                    Description = "姓名     : " + aFullName + Environment.NewLine +
                           "日期     : " + DateTime.Now.ToLocalTime().ToString() + Environment.NewLine +
                           "實收金額 : " + ((int)Convert.ToUInt32(aQryOrderPay.TSResultContent.Amount) / 100).ToString() + Environment.NewLine +
                           "付款方式 : " + "信用卡" + Environment.NewLine +
                           "說明     : " + FeeDescription + aQryOrderPay.Description + Environment.NewLine +
                           "--------------------" + Environment.NewLine;
                }
                else
                {
                    Description = "姓名     : " + aFullName + Environment.NewLine +
                           "日期     : " + DateTime.Now.ToLocalTime().ToString() + Environment.NewLine +
                           "實收金額 : " + ((int)Convert.ToUInt32(aQryOrderPay.TSResultContent.Amount) / 100).ToString() + Environment.NewLine +
                           "付款方式 : " + "ATM轉帳/匯款" + Environment.NewLine +
                           "說明     : " + FeeDescription + aQryOrderPay.Description + Environment.NewLine +
                           "--------------------" + Environment.NewLine;
                }
                #endregion
                if (aQryOrderPay.Status == "S")
                {
                    if (aQryOrderPay.TSResultContent.OrderNo.StartsWith("C"))
                    {
                        // 收費單付款日期
                        this.m_ToolUtilityClass.SetEntityDateTimeAttribute(ref aFeeEntity, "new_pay_date", DateTime.Now.ToLocalTime());
                        // 收費單實收金額
                        this.m_ToolUtilityClass.SetEntityMoneyAttribute(ref aFeeEntity, "new_fee_really_paid", new Money((int)Convert.ToUInt32(aQryOrderPay.TSResultContent.Amount) / 100));
                        // 收費單付款方式
                        this.m_ToolUtilityClass.SetOptionSetAttribute(ref aFeeEntity, "new_pay_way", 100000001); // 100000001 = 信用卡
                                                                                                                 // 收費單付款狀態
                        this.m_ToolUtilityClass.SetOptionSetAttribute(ref aFeeEntity, "new_pay_status", 100000001); // 100000001 = 信用卡已繳費
                                                                                                                    // 收費單說明
                        this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeEntity, "new_description", Description);

                        // 更新收費單
                        this.m_ToolUtilityClass.UpdateEntity(ref aFeeEntity);

                        // LINE 通知付款人
                        this.m_PushUtility.SendMessage(UserLineId, "信用卡付款結果成功!" + Environment.NewLine + Description);

                    }
                    else
                    {
                        // 收費單付款日期
                        this.m_ToolUtilityClass.SetEntityDateTimeAttribute(ref aFeeEntity, "new_pay_date", DateTime.Now.ToLocalTime());

                        // 收費單實收金額
                        this.m_ToolUtilityClass.SetEntityMoneyAttribute(ref aFeeEntity, "new_fee_really_paid", new Money((int)Convert.ToUInt32(aQryOrderPay.TSResultContent.Amount) / 100));
                        // 收費單付款方式
                        this.m_ToolUtilityClass.SetOptionSetAttribute(ref aFeeEntity, "new_pay_way", 100000002); // 100000002 = ATM轉帳/匯款
                                                                                                                 // 收費單付款狀態
                        this.m_ToolUtilityClass.SetOptionSetAttribute(ref aFeeEntity, "new_pay_status", 100000002); // 100000002 = ATM轉帳/匯款已繳費
                                                                                                                    // 收費單說明
                        this.m_ToolUtilityClass.SetEntityStringAttribute(ref aFeeEntity, "new_description", Description);

                        // 更新收費單
                        this.m_ToolUtilityClass.UpdateEntity(ref aFeeEntity);

                        // LINE 通知付款人
                        this.m_PushUtility.SendMessage(UserLineId, "ATM轉帳/匯款付款結果成功!" + Environment.NewLine + Description);
                    }
                }
                else
                {
                }

                return Json(new Dictionary<string, string>()
                {
                    { "Status", "S" }
                });
            }
            catch (System.Exception e)
            {
                String ErrorString = "ERROR : FullName = " + this.GetType().FullName.ToString() + " , Time = " + DateTime.Now.ToString() + " , Description = " + e.ToString();

                this.m_PushUtility.SendMessage( MENGSUNG_LINE_ID, ErrorString );

                throw e;
            }

        }

        private string ConvertShopNoToOrganization( String ShopNo )
        {
            switch (ShopNo)
            {
                case "NA0149_001":
                    return "yhchurchback";
                    //return "jesus";
                default:
                    return null;
            }
        }
        private string ConvertShopNoToChannelAccessToken(String ShopNo)
        {
            switch (ShopNo)
            {
                case "NA0149_001":
                    return @"HeuLkSEF5CX7hdZo4956IPpgJNdb8VqRZeL1Gu37kFFm+1F7DObAGjfeVYaggzwjZ5H4qraesvquODt7Y81jbtspNZkEq5n3oLDG+G32xQsRx1jCobkABL/Z7RKjkSACNT6h72bPQXsVn9aCuI5OogdB04t89/1O/w1cDnyilFU=";
                    //return @"g1jtWWNkjbH3OCh1cKoRvPBUkCJIygNuvV/neHXR9I4J5GBgVE85inaIaTcT4AAZ1qCuqrqJXDawrUweyBqLcX97GGokXnTRQ6MxjXAutd5Yr2FkPsZnq6kMelc/C+mqNUHaVUKFAuvTD8JvXbNmpAdB04t89/1O/w1cDnyilFU=";
                    //return @"aKS4zYeq2ZpqlLd4gslkWAyYuiC+B2f1noatF1VylPvkR2+mrvJ7mwnIIXtn2Pi117NBmNTmRZL5DO5ZMYaGCj/v9+fB6Zn9sel42Jr55PlegJdrtoSvPgm4fBso1tY/7H65+cOFDQxjqhdOU69qQAdB04t89/1O/w1cDnyilFU=";
                default:
                    return @"aKS4zYeq2ZpqlLd4gslkWAyYuiC+B2f1noatF1VylPvkR2+mrvJ7mwnIIXtn2Pi117NBmNTmRZL5DO5ZMYaGCj/v9+fB6Zn9sel42Jr55PlegJdrtoSvPgm4fBso1tY/7H65+cOFDQxjqhdOU69qQAdB04t89/1O/w1cDnyilFU=";
            }
        }

    }

}
