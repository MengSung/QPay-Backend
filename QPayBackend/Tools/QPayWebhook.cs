using Line.Messaging;
using QPayBackend.Models;
using Microsoft.AspNetCore.Mvc;
using QPay.Domain;
using System;
using System.Collections.Generic;
using ToolUtilityNameSpace;

namespace QPayBackend.Tools
{

    public class QPayAtmWebhook : Controller
    {
        #region 資料區
        private LineMessagingClient m_LineMessagingClient { get; set; }

        private PushUtility m_PushUtility { get; set; }

        private QPayProcessor m_QPayProcessor { get; set; }

        private ToolUtilityClass m_ToolUtilityClass { get; set; }

        // 胡夢嵩回傳　EXCEPTION　專用的ＩＤ
        private const String MENGSUNG_LINE_ID = @"U7638e4ed509708a3573ba6d69970583d";

        // 音訊教會-雲端除錯用
        private const String CHANNEL_ACCESS_TOKEN = @"g1jtWWNkjbH3OCh1cKoRvPBUkCJIygNuvV/neHXR9I4J5GBgVE85inaIaTcT4AAZ1qCuqrqJXDawrUweyBqLcX97GGokXnTRQ6MxjXAutd5Yr2FkPsZnq6kMelc/C+mqNUHaVUKFAuvTD8JvXbNmpAdB04t89/1O/w1cDnyilFU=";
        #endregion
        #region 初始化
        public QPayAtmWebhook()
        {
            m_QPayProcessor = new QPayProcessor();
            this.m_LineMessagingClient = new LineMessagingClient(CHANNEL_ACCESS_TOKEN);
            m_PushUtility = new PushUtility(m_LineMessagingClient);
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

        ~QPayAtmWebhook()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of
            // readability and maintainability.
            Dispose(false);
        }
        #endregion
        #region 主程式
        public JsonResult QPayBackendUrl([FromBody] BackendPostData aBackendPostData)
        {
            try
            {
                QryOrderPay aQryOrderPay = new QryOrderPay();

                //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "001-QPayAtmWebhook : "+ aBackendPostData.ShopNo +"," + aBackendPostData.PayToken);
                // 取得訂單
                aQryOrderPay = m_QPayProcessor.OrderPayQuery(aBackendPostData.ShopNo, aBackendPostData.PayToken);
                //aQryOrderPay = m_QPayProcessor.OrderPayQuery( aBackendPostData.PayToken );
                //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "001.1-QPayAtmWebhook : " + aBackendPostData.ShopNo + "," + aBackendPostData.PayToken);

                if (aBackendPostData.ShopNo != null)
                {
                    //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "002-QPayAtmWebhook");

                    if (aQryOrderPay != null)
                    {
                        //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "003-QPayAtmWebhook");

                        if (aQryOrderPay.TSResultContent.Param3 == "收費單")
                        {
                            //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "004-QPayAtmWebhook");

                            QPayFeeProcessor aQPayFeeProcessor = new QPayFeeProcessor();
                            return aQPayFeeProcessor.QPayBackendUrl(aQryOrderPay);
                        }
                        else if ( aQryOrderPay.TSResultContent.Param3 == "認獻單")
                        {
                            this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "001 認獻單" + "aBackendPostData.ShopNo=" + aBackendPostData.ShopNo+"aBackendPostData.PayToken=" + aBackendPostData.PayToken);

                            QPayDedicationBookingProcessor aQPayDedicationBookingProcessor = new QPayDedicationBookingProcessor();
                            this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "002 認獻單");

                            return aQPayDedicationBookingProcessor.QPayDedicationBookingProcessorReturnUrl(aQryOrderPay);
                            this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "003 認獻單");

                        }
                        else
                        {
                            //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "006-QPayAtmWebhook");

                            QPayFeeProcessor aQPayFeeProcessor = new QPayFeeProcessor();
                            return aQPayFeeProcessor.QPayBackendUrl(aQryOrderPay);
                        }
                    }
                    else
                    {
                        //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "007-QPayAtmWebhook");
                        return Json(new Dictionary<string, string>() { { "Status", "S" } });
                    }

                }
                else
                {
                    //this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, "008-QPayAtmWebhook");
                    return Json(new Dictionary<string, string>() { { "Status", "S" } });
                }
            }
            catch (System.Exception e)
            {
                String ErrorString = "ERROR : FullName = " + this.GetType().FullName.ToString() + " , Time = " + DateTime.Now.ToString() + " , Description = " + e.ToString();
                
                this.m_PushUtility.SendMessage(MENGSUNG_LINE_ID, ErrorString);

                //m_PushUtility.SendMessage(MENGSUNG_LINE_ID, ErrorString);
                throw e;
            }
        }
        #endregion
    }

}
