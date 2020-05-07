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
        private LineMessagingClient m_LineMessagingClient { get; }

        private PushUtility m_PushUtility { get; }

        private ReplyUtility m_ReplyUtility { get; }

        private QPayProcessor m_QPayProcessor { get; }

        ToolUtilityClass m_ToolUtilityClass;

        // 客製化
        // 永和禮拜堂 2.0
        private const String SPEECHMESSAGE_CHANNEL_ACCESS_TOKEN = @"HeuLkSEF5CX7hdZo4956IPpgJNdb8VqRZeL1Gu37kFFm+1F7DObAGjfeVYaggzwjZ5H4qraesvquODt7Y81jbtspNZkEq5n3oLDG+G32xQsRx1jCobkABL/Z7RKjkSACNT6h72bPQXsVn9aCuI5OogdB04t89/1O/w1cDnyilFU=";

        public QPayAtmWebhook()
        {
            this.m_LineMessagingClient = new LineMessagingClient(SPEECHMESSAGE_CHANNEL_ACCESS_TOKEN);

            //// 客製化
            m_PushUtility = new PushUtility(m_LineMessagingClient);
            m_ReplyUtility = new ReplyUtility(m_LineMessagingClient);

            m_QPayProcessor = new QPayProcessor(m_LineMessagingClient, m_PushUtility, m_ReplyUtility);

            m_ToolUtilityClass = new ToolUtilityClass("DYNAMICS365");
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
            //string ShopNo = aBackendPostData.ShopNo;
            //string PayToken = aBackendPostData.PayToken;

            QryOrderPay aQryOrderPay = new QryOrderPay();

            aQryOrderPay = m_QPayProcessor.OrderPayQuery(aBackendPostData.PayToken);

            EntityCollection FeeEntityCollection = m_ToolUtilityClass.RetrieveFeeByFetchXmlOrderNumber(aQryOrderPay.TSResultContent.OrderNo);
            Entity aFeeEntity;
            if (FeeEntityCollection.Entities.Count == 1)
            {
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
                return Json(new Dictionary<string, string>()
                {
                    { "Status", "S" }
                });
            }

            // 取得付款人
            Entity aContact = this.m_ToolUtilityClass.RetrieveEntity("contact", this.m_ToolUtilityClass.GetEntityLookupAttribute(aFeeEntity, "new_contact_new_fee"));
            // 取得付款人姓名
            String aFullName = this.m_ToolUtilityClass.GetEntityStringAttribute(aContact, "fullname");
            // 取得付款人 Line Id
            String UserLineId = this.m_ToolUtilityClass.GetEntityStringAttribute(aContact, "new_lineid");


            // 收費單描述說明
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
    }

}
