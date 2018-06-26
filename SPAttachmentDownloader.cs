using Microsoft.SharePoint.Client;
using NLog;
using System;
using System.IO;
using System.Net;
using System.Security;
using System.Configuration;
using Microsoft.SharePoint.Client.Search.Query;
using Microsoft.SharePoint.Client.Utilities;
using System.Collections.Generic;

namespace SPOnlineListDownloader
{
    class SPAttachmentDownloader
    {
        //You can get NLog from NuGet
        //http://www.nuget.org/packages/nlog
        private static Logger logger = NLog.LogManager.GetCurrentClassLogger();

        string LocalRootFolder = @"Path to your file";
        string folderName = "Applications";

        public void DownloadAttachments()
        {
            try
            {
                //int startListID;
                //Console.WriteLine("Enter Starting List ID");
                //if (!Int32.TryParse(Console.ReadLine(), out startListID))
                //{
                //    Console.WriteLine("Invalid ID");
                //    Console.WriteLine("Press any key to exit...");
                //    Console.ReadKey();
                //    return;
                //}

                string pass = ConfigurationManager.AppSettings["Password"];
                
                String siteUrl = "";
                String listName = "";
                SecureString Password = new SecureString();
                

                foreach (char c in pass.ToCharArray()) Password.AppendChar(c);
                Password.MakeReadOnly();

                using (ClientContext clientContext = new ClientContext(siteUrl))
                {
                    clientContext.Credentials = new SharePointOnlineCredentials("firstname.lastname@email.com", Password);
                    Console.WriteLine("Started Attachment Download " + siteUrl);
                    logger.Info("Started Attachment Download" + siteUrl);
                    //clientContext.Credentials = credentials;

                    //Get the Site Collection
                    Site oSite = clientContext.Site;
                    clientContext.Load(oSite);
                    clientContext.ExecuteQuery();

                    // Get the Web		CustomizedPageStatus	None	Microsoft.SharePoint.Client.CustomizedPageStatus

                    Web oWeb = clientContext.Web;
                    clientContext.Load(oWeb);
                    clientContext.ExecuteQuery();

                    Microsoft.SharePoint.Client.File fl = clientContext.Web.GetFileByServerRelativeUrl("");
                    clientContext.Load(fl);
                    clientContext.ExecuteQuery();

                    ClientResult<String> result = fl.ListItemAllFields.GetWOPIFrameUrl(SPWOPIFrameAction.View);
                    string url = fl.ListItemAllFields.ContentType.ToString();
                    clientContext.Load(fl.ListItemAllFields);
                    clientContext.ExecuteQuery();

                    CamlQuery query = new CamlQuery();
                    query.ViewXml = @"";

                    List oList = clientContext.Web.Lists.GetByTitle(listName);
                    clientContext.Load(oList);
                    clientContext.ExecuteQuery();

                    ListItemCollection items = oList.GetItems(query);
                    clientContext.Load(items);
                    clientContext.ExecuteQuery();


                    //Get WOPI URL
                    var query2 = new CamlQuery();
                    query.ViewXml = "...";

                    var listItems = oList.GetItems(query2);
                    clientContext.Load(listItems);
                    clientContext.ExecuteQuery();

                    var wopiUrls = new Dictionary<ListItem, ClientResult<string>>();

                    // The useWopi flag is set when WOPI URLs are desired
                  
                        foreach (var listItem in listItems)
                        {
                            wopiUrls[listItem] = listItem.GetWOPIFrameUrl(SPWOPIFrameAction.Edit);
                            clientContext.Load(listItem, item => item.Id);
                        }
                        // This query includes Id, FileSystemObjectType, DisplayName, and Modified
                        //clientContext.ExecuteQuery();
                    

                    //string relativeUrl = oWeb.ServerRelativeUrl;
                    //Folder retrievedFolder = oWeb.GetFolderByServerRelativeUrl(relativeUrl);
                    //clientContext.Load(retrievedFolder);
                    //clientContext.ExecuteQuery();

                    //CamlQuery camlQuery = new CamlQuery();
                    //camlQuery.ViewXml = @"";

                    //Use this statement if List is completely fresh and has no list items i.e no folders are present
                    //under the given list
                    if (oList.ItemCount == 0)
                    {
                        DateTime dt = DateTime.Now;
                        string date = dt.ToShortDateString().ToString();
                        string modifiedDate = date.Replace("/", "-");
                        string newFolderName = "HBApplication " + modifiedDate;

                        Folder newFolder = oList.RootFolder.Folders.Add(newFolderName);
                        clientContext.Load(newFolder);
                        clientContext.ExecuteQuery();
                    }

                    //Use this statement if the given list will always have some items i.e folder present in the list
                    foreach (ListItem listItem in items)
                    {

                        //if (Int32.Parse(listItem["ID"].ToString()) >= startListID)
                        //{

                            DateTime dt = DateTime.Now;
                            string date = dt.ToShortDateString().ToString();
                            string existingFolder = listItem["FileLeafRef"].ToString();
                            string modifiedDate = date.Replace("/", "-");
                            string folderToCheck = "HBApplication " + modifiedDate;

                            if (!existingFolder.Contains(modifiedDate))
                            {

                                string newFolderName = folderToCheck;

                                List docs = clientContext.Web.Lists.GetByTitle(listName);
                                clientContext.Load(docs, l => l.RootFolder);
                                clientContext.Load(docs.RootFolder, l => l.Folders);

                                docs.EnableFolderCreation = true;
                                docs.Update();
                                clientContext.ExecuteQuery();


                                Folder newFolder = docs.RootFolder.Folders.Add(newFolderName);
                                clientContext.Load(newFolder);
                                clientContext.ExecuteQuery();


                            }

                            else
                            {
                                Console.WriteLine("Folder already there !");
                            }



                            //string name = listItem["title"].ToString();

                            Console.WriteLine("Process Attachments for ID " +
                                  listItem["ID"].ToString());

                            Folder folder =
                                  oWeb.GetFolderByServerRelativeUrl("/CRM Documents/");

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
                        //}
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
