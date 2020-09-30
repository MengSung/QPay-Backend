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
        #region 初始化
        public QPayProcessor()
        {
        }
        #endregion
        #region 永豐金流程式區
        public QryOrderPay OrderPayQuery( String aShopNo, String aPayToken)
        {
            QryOrderPayReq orderPayQueryReq = new QryOrderPayReq()
            {
                ShopNo = aShopNo,
                PayToken = aPayToken
            };

            QryOrderPay retObj = QPayToolkit.OrderPayQuery(orderPayQueryReq, ConvertShopNoToHashCodeAndSite(aShopNo));

            return retObj;
        }
        #endregion
        #region 永豐金流工具區
        private string ConvertShopNoToHashCodeAndSite(String aShopNo)
        {
            //客製化
            switch (aShopNo)
            {
                case "DA1626_001":
                    // 永和禮拜堂(公司研發)
                    QPayToolkit._site = "https://funbiz.sinopac.com/QPay.WebAPI/api/";
                    return "D1695F439A69448F,7E460E920A184845,DEA83EFB714943F3,DC237C5C69914F0C";
                case "NA0149_001":
                    // 音訊教會 SandBox
                    QPayToolkit._site = "https://sandbox.sinopac.com/QPay.WebAPI/api/";
                    return "5E854757C751413F,D743D0EB06904837,08169D5445644513,8E52B5A180EE4399";
                default:
                    return "5E854757C751413F,D743D0EB06904837,08169D5445644513,8E52B5A180EE4399";
            }
        }
        #endregion
    }

}
