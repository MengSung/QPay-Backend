using System;
using System.Linq;
using Microsoft.Extensions.Logging;
using ToolUtilityNameSpace;
using Microsoft.Xrm.Sdk;
using Line.Messaging;
using ChurchReport.Models;
using static ChurchReport.Services.MyPayFeeTypeHelper;

namespace ChurchReport.Services
{
    /// <summary>
    /// MyPay LINE 通知發送服務
    /// 負責根據收費單類型發送對應的 LINE 通知訊息
    /// 
    /// 說明：此類別提供發送 LINE 訊息的邏輯，採同步呼叫外部推播 (LineMessagingClient)
    /// - 不改變原有行為，只加入詳細註解以便維護與閱讀
    /// - 所有例外都會向上拋出或記錄，維持原有錯誤處理行為
    /// </summary>
    public class MyPayNotificationService
    {
        // 注入的 ILogger 用於紀錄此服務內部流程與錯誤
        private readonly ILogger<MyPayNotificationService> _logger;
        // 用於建立各種通知文字內容
        private readonly MyPayMessageBuilder _messageBuilder;
        // 狀態判斷與時間解析等輔助邏輯
        private readonly MyPayStatusHelper _statusHelper;
        // 收費單類型判斷輔助器
        private readonly MyPayFeeTypeHelper _feeTypeHelper;

        // 注意：這是一個常數字串，包含 line channel access token
        // 請勿在公開 repo 或日誌中洩漏此值，生產環境建議改為從安全的設定來源讀取
        private const string LINE_CHANNEL_ACCESS_TOKEN = @"OMjL23DpFRDgphgN7JdzA7uCpv1wb4hXtsGh4FzxP8tHzeMyYOr/ry3BBqaRNJpVUhR6wPHLN4Wa4QiG5i3P5T/Y07swP5OjfCz9DKwTYC7T4mPb8x54pwtcqK1lIdgNm6skdZnu99fBsupEcbZLBAdB04t89/1O/w1cDnyilFU=";

        /// <summary>
        /// 建構函式：注入所需相依服務
        /// - logger: 用於紀錄運行時資訊與錯誤
        /// - messageBuilder: 建構 LINE 訊息內容
        /// - statusHelper: ?助解析交易狀態與時間
        /// - feeTypeHelper: ?助判斷收費單類型相關值
        /// </summary>
        public MyPayNotificationService(
            ILogger<MyPayNotificationService> logger,
            MyPayMessageBuilder messageBuilder,
            MyPayStatusHelper statusHelper,
            MyPayFeeTypeHelper feeTypeHelper)
        {
            _logger = logger;
            _messageBuilder = messageBuilder;
            _statusHelper = statusHelper;
            _feeTypeHelper = feeTypeHelper;
        }

        #region LINE 訊息發送

        /// <summary>
        /// 發送單則簡訊給指定的 lineId
        /// - 此方法為同步封裝（內部使用 Wait()），與原實作一致
        /// - 任何例外會被記錄後向上拋出，以便呼叫端處理
        /// </summary>
        /// <param name="lineId">目標 LINE userId</param>
        /// <param name="message">要發送的文字訊息</param>
        public void SendLineMessage(string lineId, string message)
        {
            try
            {
                // 建立 client 與 push helper 並發送
                var lineMessagingClient = new LineMessagingClient(LINE_CHANNEL_ACCESS_TOKEN);
                var pushUtility = new PushUtility(lineMessagingClient);

                // 將非同步操作轉為同步等待以維持原有行為
                pushUtility.SendMessage(lineId, message).Wait();

                // 記錄發送成功的簡單資訊（不要在日誌中寫入完整訊息內容或敏感資料）
                _logger.LogInformation($"SendLineMessage: 已發送 - LineId: {lineId}");
            }
            catch (Exception ex)
            {
                // 記錄錯誤並重新拋出，以便呼叫端可以採取補救措施
                _logger.LogError(ex, $"SendLineMessage: 發送失敗 - LineId: {lineId}");
                throw;
            }
        }

        #endregion

        #region 成功通知發送

        /// <summary>
        /// 發送 LINE 成功通知（使用 MyPayReturnModel）
        /// 依據收費單類型 (Dedication / Course / General) 建構不同的訊息
        /// 流程：
        /// 1. 取得 contactEntity 的 lineId，若無則直接 return
        /// 2. 解析金額（優先 actual_cost，再 cost）
        /// 3. 解析付款時間
        /// 4. 根據 feeType 建構對應訊息
        /// 5. 呼叫 SendLineMessage 發送
        /// </summary>
        public void SendLineNotificationByType(
            ToolUtilityClass utility,
            Entity feeEntity,
            MyPayReturnModel model,
            string fullName,
            FeeType feeType,
            Entity contactEntity)
        {
            try
            {
                // 若無 contactEntity，表示沒有可通知的對象，直接返回
                if (contactEntity == null) return;

                // 從 contactEntity 上取得 lineId，若為空則不發送
                string lineId = utility.GetEntityStringAttribute(contactEntity, "new_lineid");
                if (string.IsNullOrWhiteSpace(lineId)) return;

                // 解析要顯示的金額（優先使用 actual_cost）
                decimal amount = 0m;
                if (!string.IsNullOrEmpty(model.actual_cost) &&
                    decimal.TryParse(model.actual_cost, out var parsedActual))
                {
                    amount = parsedActual;
                }
                else if (!string.IsNullOrEmpty(model.cost) &&
                         decimal.TryParse(model.cost, out var parsedCost))
                {
                    amount = parsedCost;
                }

                // 解析完成時間字串為 DateTime（透過 statusHelper 處理格式）
                DateTime paymentTime = _statusHelper.ParseFinishTime(model.finishtime);

                string message;

                // 根據收費單類型建構不同的通知內容
                if (feeType == FeeType.Dedication)
                {
                    // 取得奉獻分類名稱，用於訊息呈現
                    int categoryValue = utility.GetOptionSetAttribute(feeEntity, "new_category");
                    string dedicationCategory = _feeTypeHelper.GetDedicationCategoryName(categoryValue);

                    message = _messageBuilder.BuildDedicationSuccessMessage(
                        fullName,
                        model.order_id,
                        model.uid,
                        amount,
                        dedicationCategory,
                        paymentTime
                    );
                }
                else if (feeType == FeeType.Course)
                {
                    // 取得課程名稱與排程資訊
                    string courseName = _feeTypeHelper.GetCourseName(utility, feeEntity);
                    string courseSchedule = utility.GetEntityStringAttribute(feeEntity, "new_course_schedule") ?? string.Empty;
                    string courseLocation = utility.GetEntityStringAttribute(feeEntity, "new_course_location") ?? string.Empty;

                    message = _messageBuilder.BuildCoursePaymentSuccessMessage(
                        fullName,
                        model.order_id,
                        model.uid,
                        amount,
                        courseName,
                        courseSchedule,
                        courseLocation,
                        paymentTime
                    );
                }
                else
                {
                    // 一般繳費通知，使用收費項目名稱或預設值
                    string itemName = utility.GetEntityStringAttribute(feeEntity, "new_name") ?? "繳費";

                    message = _messageBuilder.BuildGeneralPaymentSuccessMessage(
                        fullName,
                        model.order_id,
                        model.uid,
                        amount,
                        itemName,
                        paymentTime
                    );
                }

                // 實際發送（封裝在 SendLineMessage）
                SendLineMessage(lineId, message);
            }
            catch (Exception ex)
            {
                // 記錄錯誤並拋出以便外層能回應處理
                _logger.LogError(ex, $"[MyPay回傳] 發送LINE通知失敗 - OrderId: {model?.order_id}");
                throw;
            }
        }

        #endregion

        #region 失敗通知發送

        /// <summary>
        /// 發送 LINE 失敗通知（使用 MyPayReturnModel）
        /// 與成功通知類似，但包含失敗理由描述
        /// </summary>
        public void SendLineFailureNotificationByType(
            ToolUtilityClass utility,
            Entity feeEntity,
            MyPayReturnModel model,
            string fullName,
            FeeType feeType,
            Entity contactEntity)
        {
            try
            {
                if (contactEntity == null) return;

                string lineId = utility.GetEntityStringAttribute(contactEntity, "new_lineid");
                if (string.IsNullOrWhiteSpace(lineId)) return;

                // 優先使用 CRM 中的金額
                decimal amount = 0m;

                var shouldPayMoney = utility.GetEntityMoneyAttribute(feeEntity, "new_fee_shoud_pay");
                if (shouldPayMoney != null && shouldPayMoney.Value > 0)
                {
                    amount = shouldPayMoney.Value;
                }
                else if (!string.IsNullOrWhiteSpace(model.actual_cost) &&
                         decimal.TryParse(model.actual_cost, out var parsedActual))
                {
                    amount = parsedActual;
                }
                else if (!string.IsNullOrWhiteSpace(model.cost) &&
                         decimal.TryParse(model.cost, out var parsedCost))
                {
                    amount = parsedCost;
                }

                DateTime paymentTime = _statusHelper.ParseFinishTime(model.finishtime);
                string statusMessage = _statusHelper.GetPaymentStatusMessage(model.prc);

                string message;

                if (feeType == FeeType.Dedication)
                {
                    int categoryValue = utility.GetOptionSetAttribute(feeEntity, "new_category");
                    string dedicationCategory = _feeTypeHelper.GetDedicationCategoryName(categoryValue);

                    message = _messageBuilder.BuildDedicationFailureMessage(
                        fullName,
                        model.order_id,
                        model.uid,
                        amount,
                        dedicationCategory,
                        paymentTime,
                        statusMessage
                    );
                }
                else if (feeType == FeeType.Course)
                {
                    string courseName = _feeTypeHelper.GetCourseName(utility, feeEntity);
                    string courseSchedule = utility.GetEntityStringAttribute(feeEntity, "new_course_schedule") ?? string.Empty;
                    string courseLocation = utility.GetEntityStringAttribute(feeEntity, "new_course_location") ?? string.Empty;

                    message = _messageBuilder.BuildCoursePaymentFailureMessage(
                        fullName,
                        model.order_id,
                        model.uid,
                        amount,
                        courseName,
                        courseSchedule,
                        courseLocation,
                        paymentTime,
                        statusMessage
                    );
                }
                else
                {
                    string itemName = utility.GetEntityStringAttribute(feeEntity, "new_name") ?? "繳費";

                    message = _messageBuilder.BuildGeneralPaymentFailureMessage(
                        fullName,
                        model.order_id,
                        model.uid,
                        amount,
                        itemName,
                        paymentTime,
                        statusMessage
                    );
                }

                SendLineMessage(lineId, message);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"[MyPay回傳] 發送LINE失敗通知失敗 - OrderId: {model?.order_id}");
                throw;
            }
        }

        #endregion

        #region 舊版相容方法

        /// <summary>
        /// 發送付款通知（使用個別參數，舊版相容）
        /// 此方法保留以維持與舊系統的相容性，行為與 SendLineNotificationByType 類似
        /// </summary>
        public void SendPaymentNotificationByType(
            ToolUtilityClass utility,
            Entity feeEntity,
            string orderId,
            string transactionId,
            string cost,
            string fullName,
            string itemName,
            FeeType feeType,
            decimal amount,
            Entity contactEntity)
        {
            try
            {
                var contactId = utility.GetEntityLookupAttribute(feeEntity, "new_contact_new_fee");
                if (contactId == Guid.Empty) return;

                if (contactEntity == null)
                {
                    contactEntity = utility.RetrieveEntity("contact", contactId);
                }

                if (contactEntity == null) return;

                string lineId = utility.GetEntityStringAttribute(contactEntity, "new_lineid");
                if (string.IsNullOrWhiteSpace(lineId)) return;

                string message;

                if (feeType == FeeType.Dedication)
                {
                    message = _messageBuilder.BuildDedicationSuccessMessage(
                        fullName,
                        orderId,
                        transactionId,
                        amount,
                        itemName,
                        DateTime.Now
                    );
                }
                else if (feeType == FeeType.Course)
                {
                    string courseSchedule = utility.GetEntityStringAttribute(feeEntity, "new_course_schedule") ?? "";
                    string courseLocation = utility.GetEntityStringAttribute(feeEntity, "new_course_location") ?? "";

                    message = _messageBuilder.BuildCoursePaymentSuccessMessage(
                        fullName,
                        orderId,
                        transactionId,
                        amount,
                        itemName,
                        courseSchedule,
                        courseLocation,
                        DateTime.Now
                    );
                }
                else
                {
                    message = _messageBuilder.BuildGeneralPaymentSuccessMessage(
                        fullName,
                        orderId,
                        transactionId,
                        amount,
                        itemName,
                        DateTime.Now
                    );
                }

                SendLineMessage(lineId, message);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"SendNotification: 發送 LINE失敗 - OrderId: {orderId}");
            }
        }

        #endregion
    }
}
