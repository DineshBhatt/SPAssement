using System;
using System.Collections.Generic;
using System.Linq;
using  IO = System.IO;
using Microsoft.SharePoint.Client;
using System.Security;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Data;
using Eexcel = Microsoft.Office.Interop.Excel;


namespace AssementConsole
{
    class Program
    {
        private static Eexcel.Workbook MyBook = null;
        private static Eexcel.Application MyApp = null;
        private static Eexcel.Worksheet MySheet = null;
        static bool isError = true;
        public static void Main(string[] args)
        {
            //getting username and password
            Console.WriteLine("enter the Username");
            string userName = Console.ReadLine();
            Console.WriteLine("enter the password");
            SecureString password = getPassword();
            //getting site object with the url
            using (ClientContext context = new ClientContext(FieldValue.SiteURL))
            {
                context.Credentials = new SharePointOnlineCredentials(userName, password);
                Web web = context.Web;
                context.Load(web);
                context.ExecuteQuery();
                Console.WriteLine("started");
                //getting the Excel file object
                File file = context.Web.GetFileByUrl(FieldValue.ExcelURL);
                ClientResult<IO.Stream> data = file.OpenBinaryStream();
                context.Load(file);
                context.ExecuteQuery();
                //downloading the Excel file 
                var fileRef = file.ServerRelativeUrl;
                var fileInfo = File.OpenBinaryDirect(context, fileRef);
                //locating the path for excel file
                var fileName = IO.Path.Combine(@FieldValue.FilePath.Substring(0,46), file.Name);
               
                using (var fileStream = IO.File.Create(fileName))
                {
                    //storing the data to the local file
                    fileInfo.Stream.CopyTo(fileStream);
                }
                fileInfo.Dispose();
               //getting all the data from the Excel sheet and uploading it back
                ReadExcelData(context, file, data);
                Console.WriteLine(file.Name);
                Console.WriteLine(file.LinkingUrl);
                Console.WriteLine("Ends");
            }
            Console.ReadLine();
        }

        //**********************************Reading the data from the Eacel sheet************************************//
        public static void ReadExcelData(ClientContext clientContext, File fileName, ClientResult<IO.Stream> data)
        {
            
            string strErrorMsg = string.Empty;
            try
            {
                DataTable dataTable = new DataTable();
                List list = clientContext.Web.Lists.GetByTitle("Documents");
                clientContext.Load(list.RootFolder);
                clientContext.ExecuteQuery();
                using (IO.MemoryStream mStream = new IO.MemoryStream())
                {
                    if (data != null)
                    {
                        data.Value.CopyTo(mStream);
                        using (SpreadsheetDocument document = SpreadsheetDocument.Open(mStream, true))
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
                    //updatng th excel file
                    MyApp = new Eexcel.Application();
                    MyApp.Visible = false;
                    MyBook = MyApp.Workbooks.Open(FieldValue.FilePath);
                    MySheet = (Eexcel.Worksheet)MyBook.Sheets[1];
                    
                    for (int rowcount = 2; rowcount < 6; rowcount++)
                    {
                        DataRow rw = dataTable.Rows[rowcount - 2];
                        string upload = rw[0].ToString();
                        string CreatedBy = rw[2].ToString();
                        string department = rw[3].ToString();
                        string UploadStatus = rw[4].ToString();
                        string createStatus = rw[1].ToString().Substring(0,7);
                        //uploading the local files
                        IO.FileInfo fileInfo = new IO.FileInfo(upload);
                        List myList = clientContext.Web.Lists.GetByTitle("Document Libarary");
                        FileCreationInformation fileCreationInformation = new FileCreationInformation();
                        fileCreationInformation.Content = IO.File.ReadAllBytes(upload);
                        string filename = fileInfo.Name;
                        fileCreationInformation.Url = filename;
                        fileCreationInformation.Overwrite = true;
                        File fileUpload = myList.RootFolder.Files.Add(fileCreationInformation);
                        List DepartmentList = clientContext.Web.Lists.GetByTitle("Department");
                        clientContext.Load(DepartmentList);
                        clientContext.ExecuteQuery();
                        //check file size should not be greater than 15MB
                        if (fileInfo.Length < 1500000)
                        {
                            
                            clientContext.Load(fileUpload);
                            clientContext.ExecuteQuery();
                            //uploading the success status
                            try
                            {
                                User User = clientContext.Web.EnsureUser(CreatedBy); 
                                CamlQuery CamlQuery = new CamlQuery();
                                CamlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name ='Title'/><Value Type='Text'>" + department + "</Value></Eq></Where></Query><RowLimit></RowLimit></View>";
                                ListItemCollection DepartmentListItems = DepartmentList.GetItems(CamlQuery);
                                clientContext.Load(DepartmentListItems);
                                clientContext.ExecuteQuery();

                                ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                                ListItem listItem = fileUpload.ListItemAllFields;
                                clientContext.Load(listItem);
                                listItem["Create_x0020_By"] = CreatedBy;
                                listItem["Type_x0020_of_x0020_the_x0020_file"] = IO.Path.GetExtension(upload);
                                listItem["create_x0020_status"] = createStatus;
                                listItem["Department"] = DepartmentListItems[0].Id; 
                                listItem.Update();
                                clientContext.Load(listItem);
                                clientContext.ExecuteQuery();
                                MySheet.Cells[rowcount, 5] = "Succes";
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("error " + e);
                                WriteToLogs(e);
                                isError = true;
                                strErrorMsg = e.Message;
                            }
                        
                            Console.WriteLine("Data Inserted");

                        }
                        else
                        {
                            try
                            {
                                //update for failure status
                                ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                                ListItem listItem = fileUpload.ListItemAllFields;
                                clientContext.Load(listItem);
                                listItem["Create_x0020_By"] = clientContext.Web.EnsureUser(CreatedBy);
                                listItem["Type_x0020_of_x0020_the_x0020_file"] = IO.Path.GetExtension(upload);
                                listItem.Update();
                                clientContext.Load(listItem);
                                clientContext.ExecuteQuery();
                                MySheet.Cells[rowcount, 5] = "Failed";
                                MySheet.Cells[rowcount, 6] = "Size is greater than 15mb";
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("error " + e);
                                WriteToLogs(e);
                                isError = true;
                                strErrorMsg = e.Message;
                            }
                        }
                    }
                    MyBook.Save();
                    MyBook.Close();
                    //uploading the excel file
                    UploadFile(clientContext);
                }

                isError = false;
            }
            catch (Exception e)
            {
                Console.WriteLine("error " + e);
                WriteToLogs(e);
                isError = true;
                strErrorMsg = e.Message;
            }
        }
        //******************************************making the password secured***********************************************//
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
        //************************************string error to the external location************************************//
        static public void WriteToLogs(Exception e)
        {
            string Error1 = "-- " + DateTime.Now + " : " + e.StackTrace + Environment.NewLine + Environment.NewLine + Environment.NewLine;
            string Path = @"E:\logs1.txt";
            Console.WriteLine("Exists :" + IO.File.Exists(Path));
            IO.File.AppendAllText(Path, Error1);

        }
        //******************************uploading the file ****************************************************//
        public static void UploadFile(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("Documents");
            FileCreationInformation fileCreationInformation = new FileCreationInformation();
            fileCreationInformation.Content = IO.File.ReadAllBytes(FieldValue.FilePath);
            fileCreationInformation.Url = FieldValue.FilePath.Substring(47);
            fileCreationInformation.Overwrite = true;
            File fileToUpload = myList.RootFolder.Files.Add(fileCreationInformation);
            ctx.Load(fileToUpload);

            ctx.ExecuteQuery();


        }
        //*****************************************getting the cell value **********************************************//
        private static string GetCellValue(ClientContext clientContext, SpreadsheetDocument document, Cell cell)
        {
          
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
                WriteToLogs(e);
            }
            return value;
        }
    }
}
