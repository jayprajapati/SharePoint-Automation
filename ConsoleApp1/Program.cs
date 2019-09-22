using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    
    class Program
    {
        public static string[] columnsArray = { "No", "destinationpath", "status" };
        public static string Datetime = DateTime.UtcNow.ToString();
        
        static void Main(string[] args)
        {
            int line = 0;
            string CurrentDateTime = Datetime.Replace(":", "_");
            CurrentDateTime = CurrentDateTime.Replace("/", "_");
            CurrentDateTime = CurrentDateTime.Replace(" ", "_");
            string logFilePath = Environment.CurrentDirectory + "\\" + CurrentDateTime + "_MigrationTool.csv";
            //string errorLogPath = Environment.CurrentDirectory + "\\" + CurrentDateTime + "_ErrorLog.csv";

            string filename;
            Console.WriteLine("Press- \n \t 1 => Fix From Excel File \n  \t 2 => Automatic Migrataion");
            char c = Console.ReadKey().KeyChar;
            AuthenticationManager AuthManager = new AuthenticationManager();
            if (c == '1')
            {
                Console.WriteLine("Enter Destination URL :");
                string myurl = Console.ReadLine();
                ClientContext Context = new ClientContext(myurl);
                Context = AuthManager.GetWebLoginClientContext(myurl);
                //Automation.UploadLong(Context, "123.pdf", "/sites/jay2/TEST/123.pdf", false);
                Console.WriteLine("Enter File name : ");
                filename = Console.ReadLine();
                string appendText = string.Empty;
                Console.WriteLine("Fetching destination URLS..");
                using (var reader = new System.IO.StreamReader(filename))
                {
                    Console.WriteLine("File opened!");
                    while (!reader.EndOfStream)
                    {

                        line = line + 1;
                        
                        var data = reader.ReadLine().Split(',');
                        Console.WriteLine(data[0]);
                        string Documentname = data[0];
                        if(line == 1)
                        {
                            if(Documentname == "filename")
                            continue;
                            else
                            {
                                Console.WriteLine("FileNotsupported!");
                                break;
                            }
                        }
                        string source = data[1];
                        string destination = data[2];
                        destination = destination.Replace(" ","");
                        Console.WriteLine(destination);
                        //Automation.downloadDocumentLibrary(Context, Documentname, source);
                        int x=Automation.UploadLong(Context, Documentname, destination, false);
                        //Automation.UploadLong(Context ,"yeahmf.docx", "/sites/jay/Shared Documents/1/1.1/1.1.1.1/yeahmf.docx");
                        string status = string.Empty;
                        switch (x)
                        {
                            case 0:
                                status = "File already Exists!";
                                break;
                            case 1:
                                status = "Suceess";
                                break;
                            case 2:
                                status = "Success (Local file is not deleted)";
                                break;
                         
                        }
                        appendText = line.ToString() + "," + destination + "," + status + Environment.NewLine; ;
                        logExcel(columnsArray, appendText, logFilePath);
                        
                        Console.WriteLine();

                    }
                }
                Console.WriteLine("*************************************FINISHED**************************************");



            }
            else if (c == '2')
            {
                string sUrl = "https://ldce.sharepoint.com/sites/jay";
                ClientContext Sctx = new ClientContext(sUrl);
                Sctx = AuthManager.GetWebLoginClientContext(sUrl);
                Automation.myMovefun(Sctx,"TEST");
                Console.WriteLine("Completed");
            }
            else if(c=='3'){
                Console.WriteLine("Enter Source URL :");
                //string mySUrl = Console.ReadLine();
                //string mySUrl = "https://ldce.sharepoint.com/sites/jay";
                //string myDUrl = "https://ldce.sharepoint.com/sites/jay2";


                ClientContext SourceContext = new ClientContext(mySUrl);
                SourceContext = AuthManager.GetWebLoginClientContext(mySUrl);
                Console.WriteLine("Enter Destination URL :");
                //string myDUrl = Console.ReadLine();
                ClientContext DestinationContext = new ClientContext(myDUrl);
                DestinationContext = AuthManager.GetWebLoginClientContext(myDUrl);
                Console.WriteLine("Connected to both URLs.");
                Console.WriteLine("Enter DocumentLibrary Name :");
                string documentLibraryName = "TEST";//Console.ReadLine();
                Console.Clear();
                Console.WriteLine("Starting....");
                Automation.fullMigration(SourceContext,DestinationContext,documentLibraryName);
                
                //Automation.downloadDocumentLibrary(SourceContext, Documentname, source, source);
                //int x = Automation.UploadLong(Context, Documentname, destination, false);
                //Automation.UploadLong(Context ,"yeahmf.docx", "/sites/jay/Shared Documents/1/1.1/1.1.1.1/yeahmf.docx");

            }
            Console.ReadKey();

        }

        public static void logExcel(string[] columnsArray, string appendText, string Path)
        {
            string columns = string.Empty;
            for (int i = 0; i < columnsArray.Length; i++)
            {
                columns += columnsArray[i] + ",";
            }

            if (!System.IO.File.Exists(Path))
            {
                string createText = columns.Substring(0, columns.Length - 1) + Environment.NewLine;
                System.IO.File.WriteAllText(Path, createText);
            }
            System.IO.File.AppendAllText(Path, appendText);
        }



    }
}
