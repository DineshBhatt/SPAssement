using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace AssementConsole
{
    class FieldValue
    {
        public static string SiteURL = "https://acuvatehyd.sharepoint.com/teams/IJOX";
        public static string ExcelURL = "https://acuvatehyd.sharepoint.com/:x:/r/teams/IJOX/_layouts/15/Doc.aspx?sourcedoc=%7Bc25524ea-dfe1-4f61-8b88-8d02a93be877%7D&action=default&uid=%7BC25524EA-DFE1-4F61-8B88-8D02A93BE877%7D&ListItemId=6&ListId=%7B3411566D-3EE6-4CB3-9DD4-71B1E7E0AAB3%7D&odsp=1&env=prod";
        public static string FilePath = @"C:\Users\dinesh.bhatt\source\repos\assmentProc\AssementExcel.xlsx";
        public static string GetMember(string Address)
        {
            //ClientContext context = new ClientContext(SiteURL);
            //UserCollection users = context.Web.user
            return Address;
        }
    }
}
