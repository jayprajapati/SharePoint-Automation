using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace ConsoleApp1
{
    class Automation
    {
        public static int x = 0;
        public static int y = 0;
        public static int safe_flag = 0;
        public static string[] columnsArray = { "filename", "sourcepath", "destinationpath", "status" };
        public static string Datetime;
        public static string logFilePath = "MissingFiles.csv";
        public Automation()
        {
            Datetime = DateTime.UtcNow.ToString();
            string CurrentDateTime = Datetime.Replace(":", "_");
            CurrentDateTime = CurrentDateTime.Replace("/", "_");
            CurrentDateTime = CurrentDateTime.Replace(" ", "_");
            logFilePath = Environment.CurrentDirectory + "\\" + CurrentDateTime + "_MissingFiles.csv";
        }

        public static string commaCheck(string item)
        {
            string newString = item;
            if (item.Contains(","))
            {
                newString = item.Replace(",", "");
            }
            return newString;
        }

        public static void myMovefun(ClientContext ctx, string documentLibrary)
        {
            Web web = ctx.Web;
            ctx.Load(web, w => w.Title);
            ctx.ExecuteQuery();
            string webTitle = web.Title;
            List List = web.Lists.GetByTitle(documentLibrary);
            ctx.Load(List);
            ctx.ExecuteQuery();
            ctx.Load(List.RootFolder);
            ctx.Load(List.RootFolder.Folders);
            ctx.Load(List.RootFolder.Files);

            //for root files
            FileCollection files = List.RootFolder.Files;
            ctx.Load(files);
            ctx.ExecuteQuery();
            foreach (File file in files)
            {
                string SourcePath = file.ServerRelativeUrl;
                string FileName = file.Name;

                Console.WriteLine(++x + "-->" + FileName + " | " + SourcePath);
                string appendText = FileName + "," + SourcePath + ", " + SourcePath.Replace("jay", "jay2") + Environment.NewLine;

                logExcel(columnsArray, appendText, logFilePath);
            }

            //starting for folders from root  
            FolderCollection Folders = List.RootFolder.Folders;
            ctx.Load(Folders);
            ctx.ExecuteQuery();

            foreach (Folder folder in Folders)
            {
                String name = folder.Name;
                checkFolderForFolderItems(ctx, folder);
                checkFolderForFileItem(ctx, folder);
            }


        }

        public static void fullMigration(ClientContext ctx, ClientContext dctx,string documentLibrary)
        {
            Web web = ctx.Web;
            ctx.Load(web, w => w.Title);
            ctx.ExecuteQuery();
            string webTitle = web.Title;
            List List = web.Lists.GetByTitle(documentLibrary);
            ctx.Load(List);
            ctx.ExecuteQuery();
            ctx.Load(List.RootFolder);
            ctx.Load(List.RootFolder.Folders);
            ctx.Load(List.RootFolder.Files);

            //for root files
            FileCollection files = List.RootFolder.Files;
            ctx.Load(files);
            ctx.ExecuteQuery();
            foreach (File file in files)
            {
                
                
                string destinationPath = get_destinationPath(file.ServerRelativeUrl);
                Console.Clear();
                Console.WriteLine("Source Path = "+ file.ServerRelativeUrl);
                Console.WriteLine("FileName = " + file.Name);
                Console.WriteLine("DestinationPath = " + destinationPath);
                Console.WriteLine("Please Confirm the destination path by pressing 5 ");
                char c = Console.ReadKey().KeyChar;
                if (c != '5')
                    
                    return;
                else
                {
                    Console.Clear();
                    Console.WriteLine("Starting......\n");
                }

            }
            foreach (File file in files)
            {
                
                string SourcePath = file.ServerRelativeUrl;
                string FileName = file.Name;
                string destinationPath = get_destinationPath(SourcePath);
                
                FileName = FileName.Trim();
                SourcePath = SourcePath.Trim();
                Console.WriteLine(++x + ">>>" + FileName + " | " + SourcePath);
                Console.WriteLine("[*] Filesize is " + file.Length/1024 + "KB");
                //overwrite capability
                /*try
                {
                    File target_file = web.GetFileByServerRelativeUrl(destinationPath);
                    dctx.Load(target_file);
                    dctx.ExecuteQuery();


                    
                }
                catch (ServerException ex)
                {
                    if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                    {
                    }
                    throw;
                }
                */
                //main starts from here
                string downloadStatus = booldownloadDocumentLibrary(ctx, FileName, SourcePath);
                string uploadStatus = string.Empty;
                if (downloadStatus == "200")
                {
                    uploadStatus = boolUploadLong(dctx, FileName, destinationPath, false);
                }

                string appendText = FileName + "," + SourcePath + ", " + destinationPath + "," + downloadStatus + " | " + uploadStatus + Environment.NewLine;
                logExcel(columnsArray, appendText, logFilePath);
                if (downloadStatus != "200" && uploadStatus != "200")
                {
                    ++y;
                }


            }

            //starting for folders from root  
            FolderCollection Folders = List.RootFolder.Folders;
            ctx.Load(Folders);
            ctx.ExecuteQuery();

            foreach (Folder folder in Folders)
            {
                String name = folder.Name;
                checkFolderForFolderItemsWithGraphics(ctx, folder);
                checkFolderForFileItemWithGraphic(ctx, folder);
            }
            Console.WriteLine("************************Finished********************************");
            Console.WriteLine("\n\n\n\n");
            Console.WriteLine("Total Success  : "  + x  + "| Total Error : " + y);
            Console.ReadKey();




        }

        

        public static string get_destinationPath(string SourcePath)
        {
            string destinationPath = SourcePath.Substring(1);
            int index = destinationPath.IndexOf('/');
            string final = destinationPath.Substring(index + 1);
            int finalindex = final.IndexOf('/');
            string yes = final.Substring(finalindex);

            destinationPath = "/sites/jay2" + yes;// + destinationPath.Substring(0, index);
            return destinationPath.Trim();
        }
        

        public static void checkFolderForFileItem(ClientContext ctx, Folder folder)
        {
            string[] colsArray = { "No", "destinationpath", "status" };
            if (folder.Name == "Forms")
                return;

            FileCollection FileColl = folder.Files;
            ctx.Load(FileColl);
            ctx.ExecuteQuery();
            foreach (File file in FileColl)
            {
                string SourcePath = file.ServerRelativeUrl;
                //string DestPath = ConstructUrl(SourcePath, Dctx);
                string FileName = file.Name;
                string fileRef = file.ServerRelativeUrl;
                string destinationPath = get_destinationPath(fileRef);

                
                Console.WriteLine(++x + "-->" + FileName + " | " + SourcePath);
                

            }


        }

        public static void checkFolderForFileItemWithGraphic(ClientContext ctx, Folder folder)
        {
            if (folder.Name == "Forms")
                return;
            try
            {
                FileCollection FileColl = folder.Files;
                ctx.Load(FileColl);
                ctx.ExecuteQuery();
                foreach (File file in FileColl)
                {
                    Console.Clear();
                    Console.WriteLine();
                    Console.WriteLine(++x + " items Completed!");
                    Console.WriteLine("\n\n");

                    string SourcePath = file.ServerRelativeUrl;
                    //string DestPath = ConstructUrl(SourcePath, Dctx);
                    string FileName = file.Name;
                    string fileRef = file.ServerRelativeUrl;
                    string destinationPath = get_destinationPath(fileRef);

                    Console.WriteLine(">>> " + FileName + " | " + SourcePath);
                    Console.WriteLine(">>> Filesize is " + file.Length / 1024 + "KB\n");
                    string downloadStatus = booldownloadDocumentLibrary(ctx, FileName, SourcePath);
                    string uploadStatus = string.Empty;
                    if (downloadStatus == "200")
                    {
                        uploadStatus = boolUploadLong(ctx, FileName, destinationPath, false);
                    }
                    string appendText = FileName + "," + SourcePath + ", " + destinationPath + "," + downloadStatus + " | " + uploadStatus + Environment.NewLine;
                    logExcel(columnsArray, appendText, logFilePath);

                    //downloadDocumentLibrary(ctx, FileName, SourcePath);
                    //1UploadLong(ctx, FileName, destinationPath, false);
                    // string appendText = FileName + "," + SourcePath + ", " + destinationPath + Environment.NewLine;
                    // logExcel(columnsArray, appendText, logFilePath);

                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message.ToString());
            }
            


        }

        public static void checkFolderForFolderItemsWithGraphics(ClientContext ctx, Folder folder)
        {
            ++x;
            if (folder.Name == "Forms")
                return;

            FolderCollection folders = folder.Folders;
            ctx.Load(folders);
            ctx.ExecuteQuery();
            if (folders.Count == 0)
            {
                return;
            }
            foreach (Folder myfolder in folders)
            {
                checkFolderForFolderItemsWithGraphics(ctx, myfolder);
                checkFolderForFileItemWithGraphic(ctx, myfolder);
            }
        }

        public static void checkFolderForFolderItems(ClientContext ctx, Folder folder)
        {
            ++x;
            if (folder.Name == "Forms")
                return;

            FolderCollection folders = folder.Folders;
            ctx.Load(folders);
            ctx.ExecuteQuery();
            if (folders.Count == 0)
            {
                return;
            }
            foreach (Folder myfolder in folders)
            {
                checkFolderForFolderItems(ctx, myfolder);
                checkFolderForFileItem(ctx, myfolder);
            }
        }

        /// <summary>
        /// To download documentlibrary from the source.
        /// </summary>
        /// <param name="ctx">ClientContext sharepoint</param>
        /// <param name="filename">string appendText</param>
        /// <param name="filepath">string Path</param>
        /// <param name="fileref">string fileReference</param>
        /// 
        public static void downloadDocumentLibrary(ClientContext ctx, string filename, string fileref)
        {
            Console.WriteLine("[*] Downloading - " + filename);
            try
            {
                using (var fileinfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, fileref))
                {
                    filename = removeSpecialChar(filename);
                    filename = commaCheck(filename);

                    using (var filestream = System.IO.File.Create(filename))
                    {
                        fileinfo.Stream.CopyTo(filestream);
                        filestream.Close();
                    }
                    fileinfo.Dispose();
                    Console.WriteLine("[+] File Downloaded - " + filename);
                }
            }
            catch (Exception e)
            {

                Console.WriteLine("[-] Error - File Download");
                Console.WriteLine(e.ToString());
            }
            finally {
                ctx = null;
                GC.Collect();
            }
            
        }

        public static string booldownloadDocumentLibrary(ClientContext ctx, string filename, string fileref)
        {
            Console.WriteLine("\t[*] Downloading - " + filename);
            try
            {
                using (var fileinfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, fileref))
                {
                    filename = removeSpecialChar(filename);
                    filename = commaCheck(filename);

                    using (var filestream = System.IO.File.Create(filename))
                    {
                        fileinfo.Stream.CopyTo(filestream);
                        filestream.Close();
                    }
                    fileinfo.Dispose();
                    Console.WriteLine("\t[+] File Downloaded - " + filename);
                    return "200";
                }
            }
            catch (Exception e)
            {

                Console.WriteLine("\t[-] Error - File Download");
                Console.WriteLine(e.ToString());
                return e.ToString();
            }
            

        }


        public static string removeSpecialChar(string name)
        {

            //string my_String = Regex.Replace(name, @"[^0-9a-zA-Z]+", " ");
            name = name.Replace('#',' ');
            name = name.Replace('@',' ');
            name = name.Replace('!',' ');
            name = name.Replace('$',' ');
            name = name.Replace('^',' ');
            name = name.Replace('&',' ');
            name = name.Replace('*', ' ');
            name = name.Replace('_',' ');

            return name.Trim();
        }

        public static int UploadLong(ClientContext ctx, string fName, string dtn,Boolean after_del )
        {
            //Microsoft.SharePoint.Client.File.SaveBinaryDirect(ctx, "/sites/jay2/Shared Documents/2/123/456/book.xlsx", fs, true);
            try
            {
                
                dtn = dtn.Substring(0, dtn.LastIndexOf('/'));
                Console.WriteLine("[+] Expecting URL : " + dtn);
                fName = fName.Trim();
                Console.WriteLine("[*] Uploading - " + fName);
                System.IO.FileStream fs = new System.IO.FileStream(fName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                Microsoft.SharePoint.Client.File.SaveBinaryDirect(ctx, dtn + "/ " + fName, fs, false);
                Console.WriteLine("[+] Uploaded - "+fName);
                fs.Close();
                if (after_del)
                {
                    try
                    {
                        // If file found, delete it
                        if (System.IO.File.Exists(fName))
                        {
                            System.IO.File.Delete(fName);
                            Console.WriteLine("[+] File deleted.");
                            
                            return 1;
                        }
                        else
                        {
                            Console.WriteLine("[!] File not found!!!!");
                            return 2;
                        }


                    }
                    catch
                    {
                        Console.WriteLine("[-] File Not deleted!");
                        return 2;
                    }
                }
                
            }
            
            catch (Exception e)
            {
                //Console.WriteLine(e.Message.ToString());
                Console.WriteLine("[-] File Already Exists!");
                return 0;
            }
            
            return 1;
            
        }

        public static string boolUploadLong(ClientContext ctx, string fName, string dtn, Boolean after_del)
        {
            //Microsoft.SharePoint.Client.File.SaveBinaryDirect(ctx, "/sites/jay2/Shared Documents/2/123/456/book.xlsx", fs, true);
            try
            {

                dtn = dtn.Substring(0, dtn.LastIndexOf('/'));
                Console.WriteLine("\t[+] Expecting URL : " + dtn);
                fName = fName.Trim();
                Console.WriteLine("\t[*] Uploading - " + fName);
                System.IO.FileStream fs = new System.IO.FileStream(fName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                Microsoft.SharePoint.Client.File.SaveBinaryDirect(ctx, dtn + "/ " + fName, fs, true);
                Console.WriteLine("\t[+] Uploaded - " + fName);
                fs.Close();
                if (after_del)
                {
                    try
                    {
                        // If file found, delete it
                        if (System.IO.File.Exists(fName))
                        {
                            System.IO.File.Delete(fName);
                            Console.WriteLine("\t[+] File deleted.");

                            return "200";
                        }
                        else
                        {
                            Console.WriteLine("\t[!] File not found!!!!");
                            return "200";
                        }


                    }
                    catch
                    {
                        Console.WriteLine("\t[-] File Not deleted!");
                        return "300";
                    }
                }

            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message.ToString());
                Console.WriteLine("\t[-] Error Uploading - " + fName);
                return e.Message.ToString();
            }

            return "200";

        }


        public static void logExcel(string[] columnsArray, string appendText, string Path)
        {
            

            if (!System.IO.File.Exists(Path))
            {
                string columns = string.Empty;
                for (int i = 0; i < columnsArray.Length; i++)
                {
                    columns += columnsArray[i] + ",";
                }
                string createText = columns.Substring(0, columns.Length - 1) + Environment.NewLine;
                System.IO.File.WriteAllText(Path, createText);
            }

            System.IO.File.AppendAllText(Path, appendText);
        }

        


    }
}
