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
            String A21 = "";
            String A22 = "";
            String B21 = "";
            String B22 = "";

            //客製化
            switch (aShopNo)
            {
                case "NA0149_001":
                    // 音訊教會 SandBox
                    QPayToolkit._site = "https://sandbox.sinopac.com/QPay.WebAPI/api/";
                    return "5E854757C751413F,D743D0EB06904837,08169D5445644513,8E52B5A180EE4399";
                case "DA1626_001":
                    // 永和禮拜堂"板橋民族分行"
                    QPayToolkit._site = "https://funbiz.sinopac.com/QPay.WebAPI/api/";
                    return "D1695F439A69448F,7E460E920A184845,DEA83EFB714943F3,DC237C5C69914F0C";
                case "DA1626_003":
                    // 永和禮拜堂"永和分行"
                    QPayToolkit._site = "https://funbiz.sinopac.com/QPay.WebAPI/api/";
                    return "2C5D55945FCF4767,76052054D7054EA6,13F282F8A0F5475D,D782B4F1893A4334";
                case "DA2424_001":
                    // iM行動教會
                    QPayToolkit._site = "https://funbiz.sinopac.com/QPay.WebAPI/api/";
                    return "9825732578154B95,C89A75CD59D0430F,DAB73CB2A41E47FF,B09695CE58FA4774";
                case "DA2659_001":
                    // 台北得勝靈糧堂
                    QPayToolkit._site = "https://funbiz.sinopac.com/QPay.WebAPI/api/";
                    return "C8DAEA50FFB64CF4,F141E5BBE21B4D47,A922E0C106D14C35,CA22A88D1032412F";
                case "DA2890_001":
                    // 忠孝路長老教會
                    QPayToolkit._site = "https://funbiz.sinopac.com/QPay.WebAPI/api/";
                    return "BDC962CCC8AB4AE2,946D46DBDDDE43E0,6038DFB03B4342AE,B1F64046CB2E44FC";
                case "DA3033_001":
                    // 東湖禮拜堂
                    QPayToolkit._site = "https://funbiz.sinopac.com/QPay.WebAPI/api/";
                    return "4B1657DE6F3547A3,3AB478872D0A49C7,0748F400DD834C07,6506CD86B0174396";
                case "DA3190_001":
                    // 楊梅靈糧堂
                    QPayToolkit._site = "https://funbiz.sinopac.com/QPay.WebAPI/api/";
                    String A11 = "1E582BECE43F421A";
                    String A12 = "8F6ACB29B8EF4C67";
                    String B11 = "8C06D1D49C544C51";
                    String B12 = "041D9136AA9647F2";
                    return A11 + "," + A12 + "," + B11 + "," + B12;
                case "DA3189_001":
                    // 以利亞之家
                    QPayToolkit._site = "https://funbiz.sinopac.com/QPay.WebAPI/api/";
                    A21 = "A88FB80292D6420D";
                    A22 = "3844DD3B214D487C";
                    B21 = "27BC1983D2914C11";
                    B22 = "32D5A23910734C93";
                    return A21 + "," + A22 + "," + B21 + "," + B22;
                case "DA3412_001":
                    // 安平靈糧堂
                    QPayToolkit._site = "https://funbiz.sinopac.com/QPay.WebAPI/api/";
                    A21 = "2B27264C1D794727";
                    A22 = "7C91CB903482427D";
                    B21 = "7360D573A5A34184";
                    B22 = "3C85541425624385";
                    return A21 + "," + A22 + "," + B21 + "," + B22;
                case "DA3806_001":
                    // 好消息協會
                    QPayToolkit._site = "https://funbiz.sinopac.com/QPay.WebAPI/api/";
                    A21 = "81F5DAFEAFD343EC";
                    A22 = "80BA10061E59467B";
                    B21 = "B5F2CBA592004D2D";
                    B22 = "D6D805E2CF514E12";
                    return A21 + "," + A22 + "," + B21 + "," + B22;
                case "DA3855_002":
                    // 法國號靈糧堂
                    QPayToolkit._site = "https://funbiz.sinopac.com/QPay.WebAPI/api/";
                    A21 = "08B9715C313F4ABB";
                    A22 = "E8AC362AB9174D3C";
                    B21 = "81D71D28D7E04414";
                    B22 = "927ADFBE9F854C81";
                    return A21 + "," + A22 + "," + B21 + "," + B22;
                case "DA4001_001":
                    // 社團法人台灣基督教天母豐盛協會
                    QPayToolkit._site = "https://funbiz.sinopac.com/QPay.WebAPI/api/";
                    A21 = "B2FC3849C9F6487C";
                    A22 = "6ADDD7D7CCFC48BA";
                    B21 = "2F83CE17C6044E3D";
                    B22 = "48737E77D6864915";
                    return A21 + "," + A22 + "," + B21 + "," + B22;
                case "DA3009_001":
                    // 神住611靈糧堂
                    QPayToolkit._site = "https://funbiz.sinopac.com/QPay.WebAPI/api/";
                    A21 = "D3AA59886C7041B2";
                    A22 = "4519D42101984D8E";
                    B21 = "93BCEDA52A8C45D9";
                    B22 = "F983B7D4C9154484";
                    return A21 + "," + A22 + "," + B21 + "," + B22;
                case "DA4195_001":
                    // 南崁長老教會
                    QPayToolkit._site = "https://funbiz.sinopac.com/QPay.WebAPI/api/";
                    A21 = "B83DCBFA2D994F19";
                    A22 = "6ED32787DA504871";
                    B21 = "13E56D7A39AB4768";
                    B22 = "163EC08BC1624854";
                    return A21 + "," + A22 + "," + B21 + "," + B22;
                default:
                    QPayToolkit._site = "https://sandbox.sinopac.com/QPay.WebAPI/api/";
                    return "5E854757C751413F,D743D0EB06904837,08169D5445644513,8E52B5A180EE4399";
            }
        }
        #endregion
    }

}
