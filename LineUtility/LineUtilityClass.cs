using System;
using System.Collections.Generic;

using ToolUtilityNameSpace;
#region CRM 2011 reference
using Microsoft.Xrm.Sdk;
#endregion

using Line.Messaging;

namespace LineUtility
{
    public class LineUtilityClass
    {
        #region 系統參數
        //IServiceProvider m_ServiceProvider;
        //ITracingService m_TracingService;
        //IPluginExecutionContext m_Context;
        //
        //IOrganizationServiceFactory m_ServiceFactory;
        IOrganizationService m_CrmService;

        // 系統傳來的組織名稱
        String m_OrganizationName = "";

        #region Channel Access Token 設定

        // 客製化
        // 台北基督之家(雲端)的 Channel Access Token
        //private const String TPEHOC_CHANNEL_ACCESS_TOKEN = @"/iNy46gPp/ZXokg1Vr9RV/ZjodE3i7Q2o+k9nlH7l3pV8WzjAegGDduZc7gms8X5zrjSrDy2xSdNFud7JqjSDjwcTXZ6MJ/FF3NuhVg6WuXmMT34gAO7VZ0RWYrHXwAifVKpOyh2/8LiGgBpfo4ZXQdB04t89/1O/w1cDnyilFU=";
        private const String TPEHOC_CHANNEL_ACCESS_TOKEN = @"a5bB4sunKwoZGjbf0HvFnenCpiABmzIT6rGU4rQ25QAqDhxj8Wa+RwXKQN2CZVC3lSk2sZ2n5bqzCcvaa8J/DIOzUdLUUgq1wF6SIvcd0sL0uFWn0+XyaQXdii1QHvA4Lm+NU5wehU4zIhdxZaMMsAdB04t89/1O/w1cDnyilFU=";

        // 台北基督之家(公司內部開發測試)的 Channel Access Token
        //private const String TPEHOCBACK_CHANNEL_ACCESS_TOKEN = @"/iNy46gPp/ZXokg1Vr9RV/ZjodE3i7Q2o+k9nlH7l3pV8WzjAegGDduZc7gms8X5zrjSrDy2xSdNFud7JqjSDjwcTXZ6MJ/FF3NuhVg6WuXmMT34gAO7VZ0RWYrHXwAifVKpOyh2/8LiGgBpfo4ZXQdB04t89/1O/w1cDnyilFU=";
        private const String TPEHOCBACK_CHANNEL_ACCESS_TOKEN = @"a5bB4sunKwoZGjbf0HvFnenCpiABmzIT6rGU4rQ25QAqDhxj8Wa+RwXKQN2CZVC3lSk2sZ2n5bqzCcvaa8J/DIOzUdLUUgq1wF6SIvcd0sL0uFWn0+XyaQXdii1QHvA4Lm+NU5wehU4zIhdxZaMMsAdB04t89/1O/w1cDnyilFU=";

        //台中生命之道靈糧堂
        //private const String YANGMEILLC_CHANNEL_ACCESS_TOKEN = @"YTd17Eep3V5/nSaI1lxLW5vx//gOfVr21kpnpZ6RBOfvFrjhJYpvtmCIy7yxDi2tQ2cfP/6qGJ9raS72VwN7xhGjneynJHpCRrgJbz4GqMGMMEjLAcVB+hRRNCTNkMOY3rYyyN/W+/sTAx3HzzhsPgdB04t89/1O/w1cDnyilFU=";

        // 中和喜樂城靈糧堂
        // 資料庫後臺用楊梅靈糧堂的，但是LINE 訊息用自己的試用版
        private const String YANGMEILLC_CHANNEL_ACCESS_TOKEN = @"GdR6j6eh0zRtEhUIJAhfRh3whD57VurJPC0ugYwhuhsIAWD+fDnMIiJtVI9T/KSz2ciq72k3PxP+w4Qs5uVTRzZPbftBBjyLiuLV+TKlLj+gdEq//Od7xYMXDEChs0LRRoGL5QL/vxcUZXiAZ4/cQQdB04t89/1O/w1cDnyilFU=";

        // 財團法人高雄市基督教會錫安堂
        private const String KSZIONCHBACK_CHANNEL_ACCESS_TOKEN = @"4pOXVK9ujHBk/6wHN18lbsAb9usDmz3/w2WV7W16rH5OUB3XtUybWIMB7GWClWNfy7NU6o3kDCQkBDNJUfaDgCgnQAgXtspJoMBwXCOxxdD259QGPCTdlBYDilVn2x6fPqLxkouD2e2m8cbTEa6kbgdB04t89/1O/w1cDnyilFU=";

        // 思恩堂豐富教會(付費版)
        //private const String ABUNDANCE_CHANNEL_ACCESS_TOKEN = @"yvyzlpbDY4ctjVuC0vEYFDF4Gz9Ed6VR57AOmqEfRPqNSFa4tmlvgFqydqOsv8C5vOG3Ew1vPtBfZoJ7Psm69HH+oKtRA4UeMWi1EZp6j4hzhjC1ePmBRQOdfcbcGgDjJzC60Q8HAI/Err6YjFZwOwdB04t89/1O/w1cDnyilFU=";

        // 思恩堂豐富教會-公司內部研發(50人版)
        private const String ABUNDANCE_BACK_CHANNEL_ACCESS_TOKEN = "P5WXsECAIIwa5yj11jnXh0YKZrwwGFDTRT1n/9G1mr4cgkckgtcyOClWM8o0Nd6FsrUdy2DAyxAMDqclgCpdmN/zPXXl/t6NMo9Txo6qscJYJXrgc8VwEOSvjfRs+71IdPMAUyYWmnEjp00Mr6vxoAdB04t89/1O/w1cDnyilFU=";

        // 台中生命之道靈糧堂
        private const String WOL_CHANNEL_ACCESS_TOKEN = @"XofqB1wMctFHMnfJVkxGIivnfXKrRYyOzKSLrk2JtlGJz9o/esnnUf5dX4y8TRNBIrMsrEwm0Z38Zb3IISBzeokAtMD8B8oYFCtaqeeRj7RYqvBU8kDSIe6ECx+6DQfceAECSSO4vaHc+QSqeOLk4AdB04t89/1O/w1cDnyilFU=";

        #endregion

        String m_ChannelAccessToken = WOL_CHANNEL_ACCESS_TOKEN;

        LineMessagingClient m_LineMessagingClient ;

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

        ~LineUtilityClass()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of
            // readability and maintainability.
            Dispose(false);
        }
        #endregion

        ToolUtilityClass m_ToolUtilityClass;

        public LineUtilityClass( )
        {
            m_ToolUtilityClass = new ToolUtilityClass("DYNAMICS365");

            //SetupChannelAccessToken( ref aCrmService );

            m_LineMessagingClient = new LineMessagingClient(m_ChannelAccessToken);
        }

    }


}
