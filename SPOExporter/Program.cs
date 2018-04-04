//==========================================================
// SharePoint Online Exporter
// Version 1.0.0
// By Worakorn Chaichakan
//==========================================================
//#define DEBUG
using CsvHelper;
using log4net;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Threading;
using OfficeDevPnP.Core.Utilities;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace SPOExporter
{
    //Main Program
    class Program
    {
        //[DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetShortPathNameW", SetLastError = true)]
        //static extern int GetShortPathName(string pathName, System.Text.StringBuilder shortName, int cbShortName);

        //Declare an instance for log4net
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static string userName = ConfigurationManager.AppSettings["username"];
        static string password = ConfigurationManager.AppSettings["password"];
        static string SPOConfigFileName = ConfigurationManager.AppSettings["SPOConfigFileName"];
        static int delayTime = Convert.ToInt32(ConfigurationManager.AppSettings["delay"]);
        static string longPathFolder = ConfigurationManager.AppSettings["longPathFolder"];
        static string authenMode = ConfigurationManager.AppSettings["authenticationMode"];
        static bool consolePrintOut = Convert.ToBoolean(ConfigurationManager.AppSettings["consolePrintOut"]);

        static void Main(string[] args)
        {
            Log.Info("===============Sharepoint Online Exporter START===============");
            if (consolePrintOut) Console.WriteLine("===============Sharepoint Online Exporter START===============");
            Thread.Sleep(TimeSpan.FromSeconds(2)); // Sleep some secs

            IEnumerable<FileRecord> records = null;

            //string siteUrl = "https://wwww.sharepoint.com/sites/testsite/";
            //string userName = "test@gmail.com";
            //string password = "Test123";
            //string pathString = @"C:\temp\";
            //string SPOConfigFileName = "SPOSites.csv";

            //Read the configuration to get username, password and SPO config file

            SecureString securePassword = ConvertToSecureString(password);

            IEnumerable<SPOSite> SPOConfigList = GetSPOSiteConfig(SPOConfigFileName);

            Log.Info("...A number of Sharepoint Online sites in the config file is " + SPOConfigList.Count() + " ...");
            if (consolePrintOut) Console.WriteLine("...A number of Sharepoint Online sites in the config file is " + SPOConfigList.Count() + " ...");

            foreach (SPOSite ss in SPOConfigList)
            {
                string siteUrl = ss.URL;
                string pathString = ss.savedDir;
                string siteMapPath = "";

                Log.Info(">>>Stored Location: " + pathString + " <<<");
                if (consolePrintOut) Console.WriteLine(">>>Stored Location: " + pathString + " <<<");

                //OfficeDevPnP.Core.AuthenticationManager ca = new OfficeDevPnP.Core.AuthenticationManager();
                // var sss = new TokenHelper();
                // string targetRealm = TokenHelper.GetRealmFromTargetUrl(destinationSiteUri);

                ClientContext cc = null;
                switch(authenMode)
                {
                    case "SPO":
                        cc = new ClientContext(siteUrl);
                        cc.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                        Log.Info("...AUTHENTICATION MODE: SPO...");
                        if (consolePrintOut) Console.WriteLine("...AUTHENTICATION MODE: SPO...");
                        break;
                    case "WebLogin":
                    default:
                        cc = GetWebLoginClientContext(siteUrl);
                        Log.Info("...AUTHENTICATION MODE: WebLogin...");
                        if (consolePrintOut) Console.WriteLine("...AUTHENTICATION MODE: WebLogin...");
                        break;
                }

                using (ClientContext clientContext = cc)
                {
                    string dirName = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                    Web web = clientContext.Web;
                    //clientContext
             
                    try
                    {
                        clientContext.Load(web, website => website.Webs, website => website.Title);
                        clientContext.ExecuteQuery();

                        siteMapPath = dirName + "\\mappings\\" + web.Title;

                        if (!System.IO.Directory.Exists(siteMapPath))
                        {
                            System.IO.Directory.CreateDirectory(siteMapPath);
                            Log.Info("=> Folder Created: " + siteMapPath);
                            if (consolePrintOut) Console.WriteLine("=> Folder Created: " + siteMapPath);
                        }
                        else
                        {
                            if (consolePrintOut) Console.WriteLine(">>> {0} exists", siteMapPath);
                        }
                    }
                    catch (Exception e)
                    {
                        PrintError(e);
                    }

                    Log.Info("...Retrieve all Sharepoint sites under " + web.Title);
                    if (consolePrintOut) Console.WriteLine("...Retrieve all Sharepoint sites under " + web.Title);
                    List<WebDetail> webList = new List<WebDetail>();
                    webList.Add(new WebDetail { Title = web.Title, URL = siteUrl });
                    getSubWebs(siteUrl, web, clientContext, webList);

                    foreach (WebDetail w in webList)
                    {
                        Log.Info("----------Start scanning [" + w.Title + "]----------");
                        if (consolePrintOut) Console.WriteLine("----------Start scanning [" + w.Title + "]----------");
                        Log.Info("Site name: " + w.Title + " URL: " + w.URL);
                        if (consolePrintOut) Console.WriteLine("Site name: " + w.Title + " URL: " + w.URL);
                        string filePrefix = "MAP_";
                        string mappingFile = filePrefix + w.Title + ".csv";

                        string fffPath = siteMapPath + "\\" + mappingFile;


                        if (System.IO.File.Exists(fffPath))
                        {
                            Log.Info("...Read the existing file mapping");
                            if (consolePrintOut) Console.WriteLine("...Read the existing file mapping");
                            using (var sr = new StreamReader(fffPath))                           
                            {
                                using (var reader = new CsvReader(sr))
                                {
                                    records = reader.GetRecords<FileRecord>().ToList();
                                }
                            }
                        }

                        string sitePath = string.Format(@"{0}{1}\", pathString, w.Title);
                        try
                        {
                            if (!System.IO.Directory.Exists(sitePath))
                            {
                                System.IO.Directory.CreateDirectory(sitePath);
                                Log.Info("=> Folder Created: " + sitePath);
                                if (consolePrintOut) Console.WriteLine("=> Folder Created: " + sitePath);
                            }
                            else
                            {
                                if (consolePrintOut) Console.WriteLine(">>> {0} exists", sitePath);
                            }
                        }
                        catch (Exception e)
                        {
                            PrintError(e);
                        }
                        using (ClientContext ct = new ClientContext(w.URL))
                        {
                            List<FileRecord> files = new List<FileRecord>();
                            List<bool> sss = new List<bool>();

                            ct.Credentials = new SharePointOnlineCredentials(userName, securePassword);
                            var Libraries = ct.LoadQuery(ct.Web.Lists.Where(l => l.BaseTemplate == 101));
                            ct.ExecuteQuery();                           
                            foreach (List lib in Libraries)
                            {
                                try
                                {
                                    //var list = clientContext.Web.Lists.GetByTitle("Site Contents");
                                    //var list = clientContext.Web.Lists.GetByTitle("Documents");

                                    var rootFolder = lib.RootFolder;
                                    Log.Info("Document Libray: " + lib.Title + " >>>Load documents to memory...");
                                    if (consolePrintOut) Console.WriteLine("Document Libray: " + lib.Title + " >>>Load documents to memory...");
                                    ct.Load(lib);

                                    string folderPath = string.Format(@"{0}{1}\", sitePath, lib.Title);
                                    if (!System.IO.Directory.Exists(folderPath))
                                    {
                                        System.IO.Directory.CreateDirectory(folderPath);
                                        Log.Info("=> Folder Created: " + folderPath);
                                        if (consolePrintOut) Console.WriteLine("=> Folder Created: " + folderPath);
                                    }
                                    else
                                    {
                                        if (consolePrintOut) Console.WriteLine(">>> {0} exists", folderPath);
                                    }

#if DEBUG
                                    if (lib.Title == "AUS - Information Management") { continue; }
#endif                            
                                    Log.Info(">>>Checking folders and files for [" + lib.Title + "] ...");
                                    if (consolePrintOut) Console.WriteLine(">>>Checking folders and files for [" + lib.Title + "] ...");
                                    GetFoldersAndFiles(mappingFile, rootFolder, ct, folderPath, files, records, sss, folderPath);
                                }
                                catch (Exception e)
                                {
                                    PrintError(e);
                                }
                            }

                            //Validate the mapping file
                            if (sss.All(s => s == true))
                            {
                                Log.Info(">>>>>NO files changed in [" + w.Title + "] <<<<<");
                                if (consolePrintOut) Console.WriteLine(">>>>>NO files changed in [" + w.Title + "] <<<<<");
                            }
                            else
                            {
                                Log.Info("...Files changed...");
                                if (consolePrintOut) Console.WriteLine("...Files changed...");
                                if (System.IO.File.Exists(fffPath))
                                {
                                    try
                                    {
                                        System.IO.File.Delete(fffPath);
                                    }
                                    catch (Exception e)
                                    {
                                        PrintError(e);
                                    }
                                }

                                Log.Info(">>>>>TOTOL " + sss.Where(s => s == false).Count() + " File(s) Created/Updated in [" + w.Title + "] <<<<<");
                                if (consolePrintOut) Console.WriteLine(">>>>>TOTOL " + sss.Where(s => s == false).Count() + " File(s) Created/Updated in [" + w.Title + "] <<<<<");

                                using (var writer = new StreamWriter(fffPath))
                                {
                                    using (var csv = new CsvWriter(writer))
                                    {
                                        //csv.WriteHeader<FileRecord>();
                                        //csv.NextRecord();
                                        csv.WriteRecords(files); // where values implements IEnumerable
                                    }
                                    //csv.Configuration.Encod = Encoding.UTF8;
                                }
                                Log.Info("Finished writing the file mapping...");
                                if (consolePrintOut) Console.WriteLine("Finished writing the file mapping...");
                            }
                            Log.Info("----------End scanning [" + w.Title + "]----------");
                            if (consolePrintOut) Console.WriteLine("----------End scanning [" + w.Title + "]----------");
                        }
                    }

                }
            }

            //End of the process
            Log.Info("===============!!!COMPLETED!!!===============");
            if (consolePrintOut) Console.WriteLine("===============!!!COMPLETED!!!===============");
            //Console.ReadKey();
        }


        private static void getSubWebs(string path, Web myWeb, ClientContext cc, List<WebDetail> wList)
        {
            try
            {
                //ClientContext clientContext = new ClientContext(path);
                //Web oWebsite = clientContext.Web;
                cc.Load(myWeb, website => website.Webs, website => website.Title);
                cc.ExecuteQuery();
                foreach (Web orWebsite in myWeb.Webs)
                {
                    //string newpath = path + orWebsite.ServerRelativeUrl;
                    string newpath = orWebsite.Url;
                    getSubWebs(newpath, orWebsite, cc, wList);
                    wList.Add(new WebDetail { Title = orWebsite.Title, URL = newpath });
                }
            }
            catch (Exception e)
            {
                PrintError(e);
            }
        }

        private static void GetFoldersAndFiles(string mappedFile, Folder mainFolder, ClientContext clientContext, string pathString, List<FileRecord> files, IEnumerable<FileRecord> rrr, List<bool> filesExist, string libPath)
        {
            var delay = TimeSpan.FromSeconds(120);
            Thread.Sleep(TimeSpan.FromSeconds(delayTime)); // Sleep some secs
            try {
                clientContext.Load(mainFolder, k => k.Files, k => k.Folders);
                clientContext.ExecuteQuery();
            }catch(WebException wex)
            {
                var response = wex.Response as HttpWebResponse;
                if (response != null && response.StatusCode == (HttpStatusCode)429)
                {
                    //Log.Warn(string.Format("CSOM request exceeded usage limits. Sleeping for {0} before retrying.", delay));
                    //if (consolePrintOut) Console.WriteLine(string.Format("CSOM request exceeded usage limits. Sleeping for {0} before retrying.", delay));
                    //Add delay.
                    //Thread.Sleep(delay);
                    //Add to retry count and increase delay.
                    //retryAttempts++;
                    //backoffInterval = backoffInterval * 2;
                    for (int i = 0; i < 3; i++)
                    {
                        Log.Warn(string.Format("CSOM request exceeded usage limits. Sleeping for {0} before retrying.", delay));
                        if (consolePrintOut) Console.WriteLine(string.Format("CSOM request exceeded usage limits. Sleeping for {0} before retrying.", delay));
                        Thread.Sleep(delay);
                        try
                        {
                            clientContext.Load(mainFolder, k => k.Files, k => k.Folders);
                            clientContext.ExecuteQuery();
                            break;
                        }
                        catch (WebException wex2)
                        {
                            var res2 = wex2.Response as HttpWebResponse;
                            if (res2 != null && res2.StatusCode == (HttpStatusCode)429)
                            {
                                continue;
                            }
                            else
                            {
                                Log.Error(string.Format("*****{0} - {1}*****", res2.StatusCode, res2.StatusDescription));
                                if (consolePrintOut) Console.WriteLine(string.Format("*****{0} - {1}*****", res2.StatusCode, res2.StatusDescription));
                                throw;
                            }          
                        }
                    }

                }
                else
                {
                    Log.Error(string.Format("*****{0} - {1}*****", response.StatusCode, response.StatusDescription));
                    if (consolePrintOut) Console.WriteLine(string.Format("*****{0} - {1}*****", response.StatusCode, response.StatusDescription));
                    throw;
                }
            }

            foreach (var folder in mainFolder.Folders)
            {
                //Thread.Sleep(TimeSpan.FromSeconds(8));               
                if (pathString.Count() > 200)
                {
                    pathString = libPath + longPathFolder + "\\";
                }
                if (folder.Name != "Forms")
                {
                    string folderPath = string.Format(@"{0}{1}\", pathString, folder.Name);

                    try
                    {
                        if (!System.IO.Directory.Exists(folderPath))
                        {
                            System.IO.Directory.CreateDirectory(folderPath);
                            Log.Info("=> Folder Created: " + folderPath);
                            if (consolePrintOut) Console.WriteLine("=> Folder Created: " + folderPath);
                        }
                        else
                        {
                            if (consolePrintOut) Console.WriteLine(">>> {0} exists", folderPath);
                        }

                    }
                    catch (Exception e)
                    {
                        PrintError(e);
                    }
                    
                    GetFoldersAndFiles(mappedFile, folder, clientContext, folderPath, files, rrr, filesExist, libPath);
                    //if (folder.Files.AreItemsAvailable)
                    //{

                    //}else
                    //{
                    //    continue;
                    //}
                }
            }

            /*Parallel.ForEach(mainFolder.Files, (currentFile) => {
                var fileRef = currentFile.ServerRelativeUrl;

                var fileName = Path.Combine(pathString, currentFile.Name);

                if (ReadLogFile(mappedFile, fileRef, currentFile.TimeLastModified.ToString(), rrr))
                {

                    filesExist.Add(true);
                }
                else
                {
                    using (var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, fileRef))
                    {
                        using (var fileStream = System.IO.File.Create(fileName))
                        {
                            //ReadLogFile(fileRef, file.TimeLastModified.ToString());                         
                            fileInfo.Stream.CopyTo(fileStream);
                        }

                    }
                    Log.Info(">>>File Created/Updated:" + fileName);
                    if (consolePrintOut) Console.WriteLine(">>>File Created/Updated:" + fileName);
                    filesExist.Add(false);
                }

                var fff = new FileRecord();
                fff.FilePath = fileRef;
                fff.LastModifiedDate = currentFile.TimeLastModified.ToString();
                files.Add(fff);


            });*/

            foreach (var file in mainFolder.Files)
            {
                var fileRef = file.ServerRelativeUrl;
            
                var fileName = Path.Combine(pathString, file.Name);

                /*StringBuilder sb = new StringBuilder(300);
                int res = GetShortPathName(
                    fileName,
                    sb,
                    300
                );
                if (consolePrintOut) Console.WriteLine(sb.ToString());*/

                if (ReadLogFile(mappedFile, fileRef, file.TimeLastModified.ToString(), rrr))
                {

                        filesExist.Add(true);
                }
                else
                {
                    using(var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, fileRef))
                    {
                        using (var fileStream = System.IO.File.Create(fileName))
                        {
                            //ReadLogFile(fileRef, file.TimeLastModified.ToString());                         
                            fileInfo.Stream.CopyTo(fileStream);
                        }

                    }
                    Thread.Sleep(TimeSpan.FromSeconds(5)); // Sleep some secs
                    Log.Info(">>>File Created/Updated:" + fileName);
                    if (consolePrintOut) Console.WriteLine(">>>File Created/Updated:" + fileName);
                    filesExist.Add(false);
                }

                var fff = new FileRecord();
                fff.FilePath = fileRef;
                fff.LastModifiedDate = file.TimeLastModified.ToString();
                files.Add(fff);
                //files.All(s => s.LastModifiedDate == "");
            }
        //return files;
        }

        private static SecureString ConvertToSecureString(string password)
        {
            if (password == null)
            {
                throw new ArgumentNullException("password");
            }
            var securePassword = new SecureString();

            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }

            securePassword.MakeReadOnly();
            return securePassword;
        }

        private static bool ReadLogFile(string mappedFile, string fileRef, string modDate, IEnumerable<FileRecord> records)
        {
            //IEnumerable<FileRecord> records;
            //string fileName = "testCSV10.csv";
            string dirName = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string filePath = dirName + "\\" + mappedFile;
            if (records != null)
            {

                var bbb = records;
                foreach (var rec in records)
                {
                    if (rec.FilePath == fileRef && rec.LastModifiedDate == modDate)
                    {
                        return true;
                    }
                    
                }
            }
            else{
            }
            return false;
        }

        private static IEnumerable<SPOSite> GetSPOSiteConfig(string configFile)
        {
            IEnumerable<SPOSite> sites;
            string dirName = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string configPath = dirName + "\\" + configFile;

            if (System.IO.File.Exists(configPath))
            {
                Log.Info("...Load SPO Site Config File => " + configFile);
                if (consolePrintOut) Console.WriteLine("...Load SPO Site Config File => " + configFile);
                using (var sr = new StreamReader(configPath))                
                {
                    using (var reader = new CsvReader(sr))
                    {
                        reader.Configuration.HasHeaderRecord = false;
                        sites = reader.GetRecords<SPOSite>().ToList();
                    }
                    //var reader = new CsvReader(sr);
                }
            }
            else {
                return null;
            }
            return sites;
        }

        private static void PrintError(Exception e) {
            Log.Error("*****" + e.Message + "*****");
            if (consolePrintOut) Console.WriteLine("*****" + e.Message + "*****");
            if (e.InnerException != null)
            {
                Log.Error("*****Inner exception: {0}*****", e.InnerException);
                if (consolePrintOut) Console.WriteLine("*****Inner exception: {0}*****", e.InnerException);
            }
        }

        private static ClientContext GetWebLoginClientContext(string siteUrl, System.Drawing.Icon icon = null)
        {
            var authCookiesContainer = new CookieContainer();
            var siteUri = new Uri(siteUrl);

            var thread = new Thread(() =>
            {
                var form = new System.Windows.Forms.Form();
                if (icon != null)
                {
                    form.Icon = icon;
                }
                var browser = new System.Windows.Forms.WebBrowser
                {
                    ScriptErrorsSuppressed = true,
                    Dock = DockStyle.Fill
                };

                form.SuspendLayout();
                form.Width = 900;
                form.Height = 500;
                form.Text = $"Log in to {siteUrl}";
                form.Controls.Add(browser);
                form.ResumeLayout(false);

                browser.Navigate(siteUri);

                browser.Navigated += (sender, args) =>
                {
                    if (siteUri.Host.Equals(args.Url.Host))
                    {
                        var cookieString = CookieReader.GetCookie(siteUrl).Replace("; ", ",").Replace(";", ",");

                        // Get FedAuth and rtFa cookies issued by ADFS when accessing claims aware applications.
                        // - or get the EdgeAccessCookie issued by the Web Application Proxy (WAP) when accessing non-claims aware applications (Kerberos).
                        IEnumerable<string> authCookies = null;
                        if (Regex.IsMatch(cookieString, "FedAuth", RegexOptions.IgnoreCase))
                        {
                            authCookies = cookieString.Split(',').Where(c => c.StartsWith("FedAuth", StringComparison.InvariantCultureIgnoreCase) || c.StartsWith("rtFa", StringComparison.InvariantCultureIgnoreCase));
                        }
                        else if (Regex.IsMatch(cookieString, "EdgeAccessCookie", RegexOptions.IgnoreCase))
                        {
                            authCookies = cookieString.Split(',').Where(c => c.StartsWith("EdgeAccessCookie", StringComparison.InvariantCultureIgnoreCase));
                        }
                        if (authCookies != null)
                        {
                            authCookiesContainer.SetCookies(siteUri, string.Join(",", authCookies));
                            form.Close();
                        }
                    }
                };

                form.Focus();
                form.ShowDialog();
                browser.Dispose();
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (authCookiesContainer.Count > 0)
            {
                var ctx = new ClientContext(siteUrl);
                ctx.ExecutingWebRequest += (sender, e) => e.WebRequestExecutor.WebRequest.CookieContainer = authCookiesContainer;
                return ctx;
            }

            return null;
        }

    }
}
