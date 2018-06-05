using Microsoft.SharePoint.Client;
using NLog;
using System;
using System.IO;
using System.Net;
using System.Security;
using System.Configuration;

namespace SPOnlineListDownloader
{
    class SPAttachmentDownloader
    {
        //You can get NLog from NuGet
        //http://www.nuget.org/packages/nlog
        private static Logger logger = NLog.LogManager.GetCurrentClassLogger();

        string listName = "Custom List";
        string LocalRootFolder = @"C:\Users\vinn\Documents\sharepointdocs";

        public void DownloadAttachments()
        {
            try
            {
                int startListID;
                Console.WriteLine("Enter Starting List ID");
                if (!Int32.TryParse(Console.ReadLine(), out startListID))
                {
                    Console.WriteLine("Invalid ID");
                    Console.WriteLine("Press any key to exit...");
                    Console.ReadKey();
                    return;
                }

                string pass = ConfigurationManager.AppSettings["Password"];
                
                String siteUrl = "https://tridentcrm1.sharepoint.com";
                String listName = "Custom List";
                SecureString Password = new SecureString();
               

                foreach (char c in pass.ToCharArray()) Password.AppendChar(c);
                Password.MakeReadOnly();

                using (ClientContext clientContext = new ClientContext(siteUrl))
                {
                    clientContext.Credentials = new SharePointOnlineCredentials("vj@tridentcrm1.onmicrosoft.com", Password);
                    Console.WriteLine("Started Attachment Download " + siteUrl);
                    logger.Info("Started Attachment Download" + siteUrl);
                    //clientContext.Credentials = credentials;

                    //Get the Site Collection
                    Site oSite = clientContext.Site;
                    clientContext.Load(oSite);
                    clientContext.ExecuteQuery();

                    // Get the Web
                    Web oWeb = clientContext.Web;
                    clientContext.Load(oWeb);
                    clientContext.ExecuteQuery();

                    CamlQuery query = new CamlQuery();
                    query.ViewXml = @"";

                    List oList = clientContext.Web.Lists.GetByTitle(listName);
                    clientContext.Load(oList);
                    clientContext.ExecuteQuery();

                    ListItemCollection items = oList.GetItems(query);
                    clientContext.Load(items);
                    clientContext.ExecuteQuery();

                    foreach (ListItem listItem in items)
                    {
                        if (Int32.Parse(listItem["ID"].ToString()) >= startListID)
                        {

                            Console.WriteLine("Process Attachments for ID " +
                                  listItem["ID"].ToString());

                            Folder folder =
                                  oWeb.GetFolderByServerRelativeUrl(oSite.Url +
                                  "/Lists/" + listName + "/Attachments/" +

                                  listItem["ID"]);

                            clientContext.Load(folder);

                            try
                            {
                                clientContext.ExecuteQuery();
                            }
                            catch (ServerException ex)
                            {
                                logger.Info(ex.Message);
                                Console.WriteLine(ex.Message);
                                logger.Info("No Attachment for ID " + listItem["ID"].ToString());
                                Console.WriteLine("No Attachment for ID " + listItem["ID"].ToString());
                            }

                            FileCollection attachments = folder.Files;
                            clientContext.Load(attachments);
                            clientContext.ExecuteQuery();

                            foreach (Microsoft.SharePoint.Client.File oFile in folder.Files)
                            {
                                logger.Info("Found Attachment for ID " +
                                      listItem["ID"].ToString());

                                Console.WriteLine("Found Attachment for ID " +
                                      listItem["ID"].ToString());

                                FileInfo myFileinfo = new FileInfo(oFile.Name);
                                WebClient client1 = new WebClient();
                                //client1.Credentials = credentials;

                                logger.Info("Downloading " +
                                      oFile.ServerRelativeUrl);

                                Console.WriteLine("Downloading " +
                                      oFile.ServerRelativeUrl);

                                string localFilePath = LocalFileLocation(oFile.ServerRelativeUrl, true);
                                //string localFilePath = @"C:\Users\vinn\Documents\sharepointdocs";

                                if (!System.IO.File.Exists(localFilePath))
                                {
                                    Microsoft.SharePoint.Client.FileInformation f = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, oFile.ServerRelativeUrl);

                                    using (var fileStream = new FileStream(localFilePath, FileMode.Create))
                                    {
                                        f.Stream.CopyTo(fileStream);
                                    }
                                    Console.WriteLine("Completed!");
                                }
                                else
                                {
                                    Console.WriteLine(" already existed!");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                logger.ErrorException(e.Message, e);
                logger.Error(e.StackTrace);
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }

        }

        private string LocalFileLocation(string serverRelativeUrl, bool isFile)
        {
            //string LocalRootFolder = string.Empty;
            
            string localFilePath = Path.Combine(LocalRootFolder, serverRelativeUrl.Replace('/', Path.DirectorySeparatorChar).TrimStart('\\'));
            string directoryPath = localFilePath;
            if (isFile)
            {
                directoryPath = Path.GetDirectoryName(localFilePath);
            }
            Directory.CreateDirectory(directoryPath);
            return localFilePath;
        }

    }
}
