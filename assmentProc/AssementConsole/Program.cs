using System;
using System.Security;
using System.Data;
using Microsoft.SharePoint.Client;
using OfficeOpenXml.Core;
using OfficeOpenXml.Core.ExcelPackage;

namespace AssementConsole
{
    class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("enter the password");
            string userName = "dinesh.bhatt@acuvate.com";
            SecureString password = getPassword();
            using (var context = new ClientContext("https://acuvatehyd.sharepoint.com/teams/IJOX"))
            {
                context.Credentials = new SharePointOnlineCredentials(userName, password);
                Web web = context.Web;
                context.Load(web);
                context.ExecuteQuery();

                Console.WriteLine("doing");

                File file = context.Web.GetFileByUrl("https://acuvatehyd.sharepoint.com/:x:/r/teams/IJOX/_layouts/15/Doc.aspx?sourcedoc=%7Bc25524ea-dfe1-4f61-8b88-8d02a93be877%7D&action=default&uid=%7BC25524EA-DFE1-4F61-8B88-8D02A93BE877%7D&ListItemId=6&ListId=%7B3411566D-3EE6-4CB3-9DD4-71B1E7E0AAB3%7D&odsp=1&env=prod");
                ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                DataTable dataTable = new DataTable();
                
                    context.Load(file);
                context.ExecuteQuery();
                Console.WriteLine(file.Name);
                Console.WriteLine(file.LinkingUrl);
                Console.WriteLine("doing");


                Console.ReadLine();
            }

        }
        public static SecureString getPassword()
        {
            ConsoleKeyInfo info;
            SecureString Password = new SecureString();
            do
            {
                info = Console.ReadKey(true);
                if (info.Key != ConsoleKey.Enter)
                {
                    Password.AppendChar(info.KeyChar);
                }
            }
            while (info.Key != ConsoleKey.Enter);
            return Password;

        }
    }
}
