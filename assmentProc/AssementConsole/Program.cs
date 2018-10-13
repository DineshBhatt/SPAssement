using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using System.Security;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Data;
using System.Configuration;
using System.Net.Mail;
using Microsoft.SharePoint.Client.Utilities;


namespace AssementConsole
{
    class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("enter the password");
            string userName = "dinesh.bhatt@acuvate.com";
            SecureString password = getPassword();
            using (ClientContext context = new ClientContext("https://acuvatehyd.sharepoint.com/teams/IJOX"))
            {
                context.Credentials = new SharePointOnlineCredentials(userName, password);
                Web web = context.Web;
                context.Load(web);
                context.ExecuteQuery();
  

                Console.WriteLine("doing");

                File file = context.Web.GetFileByUrl("https://acuvatehyd.sharepoint.com/:x:/r/teams/IJOX/_layouts/15/Doc.aspx?sourcedoc=%7Bc25524ea-dfe1-4f61-8b88-8d02a93be877%7D&action=default&uid=%7BC25524EA-DFE1-4F61-8B88-8D02A93BE877%7D&ListItemId=6&ListId=%7B3411566D-3EE6-4CB3-9DD4-71B1E7E0AAB3%7D&odsp=1&env=prod");
               
                context.Load(file);
                context.ExecuteQuery();
                ReadExcelData(context, file.Name);


                Console.WriteLine(file.Name);
                Console.WriteLine(file.LinkingUrl);
                Console.WriteLine("doing");


                Console.ReadLine();
            }

        }
        public static void ReadExcelData(ClientContext clientContext, string fileName)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            try
            {
                DataTable dataTable = new DataTable();
                List list = clientContext.Web.Lists.GetByTitle("Documents");
                clientContext.Load(list.RootFolder);
                clientContext.ExecuteQuery();
                string fileServerRelativeUrl = list.RootFolder.ServerRelativeUrl + "/" + fileName;
                File file = clientContext.Web.GetFileByServerRelativeUrl(fileServerRelativeUrl);
                ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                clientContext.Load(file);
                clientContext.ExecuteQuery();
                using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                {
                    if (data != null)
                    {
                        data.Value.CopyTo(mStream);
                        using (SpreadsheetDocument document = SpreadsheetDocument.Open(mStream, false))
                        {
                            WorkbookPart workbookPart = document.WorkbookPart;
                            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                            string relationshipId = sheets.First().Id.Value;
                            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                            Worksheet workSheet = worksheetPart.Worksheet;
                            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                            IEnumerable<Row> rows = sheetData.Descendants<Row>();
                            foreach (Cell cell in rows.ElementAt(0))
                            {
                                string str = GetCellValue(clientContext, document, cell);
                                dataTable.Columns.Add(str);
                            }
                            foreach (Row row in rows)
                            {
                                if (row != null)
                                {
                                    DataRow dataRow = dataTable.NewRow();
                                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                                    {
                                        dataRow[i] = GetCellValue(clientContext, document, row.Descendants<Cell>().ElementAt(i));
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                            dataTable.Rows.RemoveAt(0);
                        }
                    }

                    for(int rowcount = 0; rowcount < 4; rowcount++)
                    {
                        string upload = string.Empty;
                        string CreatedBy = string.Empty;
                        DataRow rw = dataTable.Rows[rowcount];
                        for (int colcount = 0; colcount < 7; colcount++)
                        {
                            if (colcount == 2)
                                CreatedBy = rw[colcount].ToString();
                            
                            if (colcount == 6)
                                upload = rw[colcount].ToString();
                            
                            Console.WriteLine(rw[colcount].ToString());
                        }
                        UploadFile(clientContext, upload, CreatedBy);
                    }
                }
            
                isError = false;
            }
            catch (Exception e)
            {
                Console.WriteLine("error "+e);

                isError = true;
                strErrorMsg = e.Message;
            }
            finally
            {
                if (isError)
                {
                    //Logging
                }
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

        public static void UploadFile(ClientContext ctx, string name, string createdBy)
        {
            DocumentLibraryInformation myList = ctx.Web.GetByTitle("Document Libarary");
            FileCreationInformation fileCreationInformation = new FileCreationInformation();
            fileCreationInformation.Content = System.IO.File.ReadAllBytes(@"C:\Users\dinesh.bhatt\source\repos\assmentProc\AssementConsole\"+name);
            fileCreationInformation.Url = @"Document Libarary\"+name;
            fileCreationInformation.Overwrite = true;
            File fileUpload = myList.RootFolder.Files.Add(fileCreationInformation);
            ctx.Load(fileUpload);
            ctx.ExecuteQuery();
            ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
            ListItem listItem = myList.AddItem(listItemCreationInformation);
            //ListItem listItem = fileUpload.ListItemAllFields;
            listItem["Create_x0020_By"] = createdBy;
            listItem["Type of the file"] = fileUpload.GetType().ToString();
            //listItem["Size of the File"] = fileUpload
            ctx.Load(listItem);
            ctx.ExecuteQuery();
            Console.WriteLine("Data Inserted");

        }

        private static string GetCellValue(ClientContext clientContext, SpreadsheetDocument document, Cell cell)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            string value = string.Empty;
            try
            {
                if (cell != null)
                {
                    SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
                    if (cell.CellValue != null)
                    {
                        value = cell.CellValue.InnerXml;
                        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                        {
                            if (stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)] != null)
                            {
                                isError = false;
                                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                            }
                        }
                        else
                        {
                            isError = false;
                            return value;
                        }
                    }
                }
                isError = false;
                return string.Empty;
            }
            catch (Exception e)
            {
                isError = true;
                strErrorMsg = e.Message;
            }
            finally
            {
                if (isError)
                {
                    //Logging
                }
            }
            return value;
        }
    }
}
