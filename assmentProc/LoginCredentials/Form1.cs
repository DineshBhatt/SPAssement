using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Security;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using System.IO;
using System.Data;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Windows.Forms;

namespace LoginCredentials
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        static bool isError = false;
        public Form1()
        {
            InitializeComponent();
        }
        string userName = string.Empty;
        SecureString password;
        private void button1_Click(object sender, EventArgs e)
        {
            string userName = textBox1.Text;
         
            foreach(char c in textBox2.Text ) {
                password.AppendChar(c);
            }


            using (ClientContext context = new ClientContext("https://acuvatehyd.sharepoint.com/teams/IJOX"))
            {
                try
                {
                    context.Credentials = new SharePointOnlineCredentials(userName, password);
                    Web web = context.Web;
                    context.Load(web);
                    context.ExecuteQuery();
                }catch (IdcrlException ex)
                {
                    label3.Text = "*incorrect Password";
                   
                    throw ex;
                    
                }
                try { 
                Microsoft.SharePoint.Client.File file = context.Web.GetFileByUrl("https://acuvatehyd.sharepoint.com/:x:/r/teams/IJOX/_layouts/15/Doc.aspx?sourcedoc=%7Bc25524ea-dfe1-4f61-8b88-8d02a93be877%7D&action=default&uid=%7BC25524EA-DFE1-4F61-8B88-8D02A93BE877%7D&ListItemId=6&ListId=%7B3411566D-3EE6-4CB3-9DD4-71B1E7E0AAB3%7D&odsp=1&env=prod");
                context.Load(file);
                context.ExecuteQuery();
               

                var fileRef = file.ServerRelativeUrl;
                var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, fileRef);

                var fileName = Path.Combine(@"C:\Users\dinesh.bhatt\source\repos\assmentProc", file.Name);

                using (var fileStream = System.IO.File.Create(fileName))
                {
                    fileInfo.Stream.CopyTo(fileStream);
                }
                fileInfo.Dispose();
                ReadExcelData(context, file);
                }
                catch (IdcrlException ex)
                {
                    label3.Text = "*incorrect Password";
                    WriteToLogs(ex);
                }

            }
            Process.Start("chrome", @"https://acuvatehyd.sharepoint.com/teams/IJOX");
            Environment.Exit(0); 
        }
        private void Form1_Load(object sender, EventArgs e)
        {
          
        }
        public static void ReadExcelData(ClientContext clientContext, Microsoft.SharePoint.Client.File fileName)
        {
            bool isError = true;
            string strErrorMsg = string.Empty;
            try
            {
                System.Data.DataTable dataTable = new System.Data.DataTable();
                List list = clientContext.Web.Lists.GetByTitle("Documents");
                clientContext.Load(list.RootFolder);
                clientContext.ExecuteQuery();
                string fileServerRelativeUrl = list.RootFolder.ServerRelativeUrl + "/" + fileName.Name;
                Microsoft.SharePoint.Client.File file = clientContext.Web.GetFileByServerRelativeUrl(fileServerRelativeUrl);
                ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                clientContext.Load(file);
                clientContext.ExecuteQuery();
                using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                {
                    if (data != null)
                    {
                        data.Value.CopyTo(mStream);
                        using (SpreadsheetDocument document = SpreadsheetDocument.Open(mStream, true))
                        {
                            WorkbookPart workbookPart = document.WorkbookPart;
                            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>().Elements<Sheet>();
                            string relationshipId = sheets.First().Id.Value;
                            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                            DocumentFormat.OpenXml.Spreadsheet.Worksheet workSheet = worksheetPart.Worksheet;
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
                    MyApp = new Excel.Application();
                    MyApp.Visible = false;
                    MyBook = MyApp.Workbooks.Open(@"C:\Users\dinesh.bhatt\source\repos\assmentProc\AssementExcel.xlsx");
                    MySheet = (Excel.Worksheet)MyBook.Sheets[1];
                    for (int rowcount = 2; rowcount < 6; rowcount++)
                    {

                        string upload = string.Empty;
                        string CreatedBy = string.Empty;
                        string UploadStatus = string.Empty;
                        string createStatus = string.Empty;
                        DataRow rw = dataTable.Rows[rowcount - 2];
                        for (int colcount = 0; colcount < 6; colcount++)
                        {
                            if (colcount == 1)
                                createStatus = rw[colcount].ToString();
                            if (colcount == 2)
                                CreatedBy = rw[colcount].ToString();
                            if (colcount == 4)
                                UploadStatus = rw[colcount].ToString();
                            if (colcount == 0)
                                upload = rw[colcount].ToString();
                            Console.WriteLine(rw[colcount].ToString());
                        }
                       
                        FileInfo fileInfo = new FileInfo(upload);
                        List myList = clientContext.Web.Lists.GetByTitle("Document Libarary");
                        FileCreationInformation fileCreationInformation = new FileCreationInformation();
                        fileCreationInformation.Content = System.IO.File.ReadAllBytes(upload);
                        string filename = fileInfo.Name;

                        fileCreationInformation.Url = filename;
                        fileCreationInformation.Overwrite = true;
                        Microsoft.SharePoint.Client.File fileUpload = myList.RootFolder.Files.Add(fileCreationInformation);
                      
                        if (fileInfo.Length < 1500000)
                        {
                        
                            clientContext.Load(fileUpload);
                            clientContext.ExecuteQuery();
                            ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                            ListItem listItem = fileUpload.ListItemAllFields;
                            clientContext.Load(listItem);
                            listItem["Create_x0020_By"] = CreatedBy;
                            listItem["Type_x0020_of_x0020_the_x0020_file"] = Path.GetExtension(upload);
                            listItem["create_x0020_status"] = createStatus;
                            listItem.Update();
                            clientContext.Load(listItem);
                            clientContext.ExecuteQuery();
                            MySheet.Cells[rowcount, 5] = "Succes";

                            Console.WriteLine("Data Inserted");

                        }
                        else
                        {
                            ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                            ListItem listItem = fileUpload.ListItemAllFields;
                            clientContext.Load(listItem);
                            listItem["Create_x0020_By"] = CreatedBy;
                            listItem["Type_x0020_of_x0020_the_x0020_file"] = Path.GetExtension(upload);
                            listItem.Update();
                            clientContext.Load(listItem);
                            clientContext.ExecuteQuery();
                            MySheet.Cells[rowcount, 5] = "Failed";
                            MySheet.Cells[rowcount, 6] = "Size is greater than 15mb";
                        }
                    }
                    MyBook.Save();
                    MyBook.Close();
                    UploadFile(clientContext);
                }

                isError = false;
            }
            catch (Exception e)
            {
                Console.WriteLine("error " + e);

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
      
        //******************************uploading the file********************************************//
        public static void UploadFile(ClientContext ctx)
        {
            List myList = ctx.Web.Lists.GetByTitle("Documents");
            FileCreationInformation fileCreationInformation = new FileCreationInformation();
            fileCreationInformation.Content = System.IO.File.ReadAllBytes(@"C:\Users\dinesh.bhatt\source\repos\assmentProc\AssementExcel.xlsx");
            fileCreationInformation.Url = "AssementExcel.xlsx";
            fileCreationInformation.Overwrite = true;
            Microsoft.SharePoint.Client.File fileToUpload = myList.RootFolder.Files.Add(fileCreationInformation);
            ctx.Load(fileToUpload);

            ctx.ExecuteQuery();


        }
        //******************************getting the cell value********************************************//
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
            }
            return value;
        }
       

        private void button2_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }
        public static void WriteToLogs(Exception e)
        {
            string Error1 = "-- " + DateTime.Now + " : " + e.StackTrace + Environment.NewLine + Environment.NewLine + Environment.NewLine;
            string Path = @"D:\logs1.txt";
            //var myfile = File.Create(path);

            Console.WriteLine("Exists :" + System.IO.File.Exists(Path));
            System.IO.File.AppendAllText(Path, Error1);

        }
        private void ErrorFoundApps_Check(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }
    }
}
