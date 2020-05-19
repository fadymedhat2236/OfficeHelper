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
        public static readonly String ServiceFlow = "Outbound";

        //for the documents properties
        public static readonly String Subject = "B2B REDF Subsidy Schedules Upload";
        public static readonly String PublishDate = "18-May-2020";
        public static readonly String ReviewDate = "19-May-2020";
        public static readonly String SSLClientCryptoProfile = "elmGW_ssl_client_profile";
        public static readonly String XMLManager = "elmGW_xml_mngr";
        public static readonly String ServiceCanonicalName = "B2BREDFSubsidySchedsUpload";
        public static readonly String ServiceID  = "142";
        public static readonly String ServiceSubCategory = "REDF";
        public static readonly String BackendName = "REDF";

        //folder paths
        public static readonly String TemplatesFolderPath = "C:\\Users\\fmedhat\\source\\repos\\OfficeHelper";
        public static readonly String TFSFolderPath = "D:\\Projects\\SABB_TFS\\(Common)\\Documents\\Design\\Services\\Gateway_DP";
        
        //public static readonly String TFSFolderPath = "D:\\Test";

    }
}
