using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeHelper
{
    class Constants
    {
        //service flow
        public static readonly String ServiceFlow = "Inbound";

        //for the documents properties
        public static readonly String Subject = "B2B Report Download";
        public static readonly String PublishDate = "14-Jan-2019";
        public static readonly String ReviewDate = "17-Jan-2019";
        public static readonly String SSLClientCryptoProfile = "elmGW_ssl_client_profile";
        public static readonly String XMLManager = "elmGW_xml_mngr";
        public static readonly String ServiceCanonicalName = "B2BReportDownload";
        public static readonly String ServiceID  = "126";
        public static readonly String ServiceSubCategory = "Musaned";
        public static readonly String BackendName = "ESB";

        //folder paths
        public static readonly String TemplatesFolderPath = "C:\\Users\\fmedhat\\source\\repos\\OfficeHelper";
        //public static readonly String TFSFolderPath = "D:\\Projects\\SABB_TFS\\(Common)\\Documents\\Design\\Services\\Gateway_DP";
        
        public static readonly String TFSFolderPath = "D:\\Test";

    }
}
