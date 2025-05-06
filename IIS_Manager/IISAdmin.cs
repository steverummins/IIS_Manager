using Microsoft.Web.Administration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.DirectoryServices;
using System.IO;
using System.Linq;
using System.Management;
using System.Management.Instrumentation;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Xml;

namespace IIS_Manager
{
    /// <summary>
    /// Provides a collection of utility methods for managing IIS websites,
    /// application pools, FTP directories, virtual applications, and user permissions.
    /// Designed for automation and administrative scripting scenarios.
    /// </summary>

    public class IISAdmin
    {
        /// <summary>
        /// Assigns an existing application pool to the root application of a specified IIS website.
        /// </summary>
        /// <param name="sitename">The name of the IIS site.</param>
        /// <param name="appPool">The name of the application pool to assign.</param>

        public static void AssignAppPooltoSite(string sitename, string appPool)
        {
            Microsoft.Web.Administration.ServerManager manager = new Microsoft.Web.Administration.ServerManager();
            Site defaultSite = manager.Sites[sitename];
            // defaultSite.Applications will give you the list of 'this' web site reference and all
            // virtual directories inside it -- 0 index is web site itself.
            Microsoft.Web.Administration.Application oVDir = defaultSite.Applications["/"];
            oVDir.ApplicationPoolName = appPool;
            manager.CommitChanges();
        }

        /// <summary>
        /// Creates a new local Windows user and adds it to the "Guests" group.
        /// </summary>
        /// <param name="username">The username of the new account.</param>
        /// <param name="password">The password for the account.</param>
        /// <param name="description">A description for the user account.</param>

        public static void CreateUser(string username, string password, string description)
        {
            try
            {
                DirectoryEntry AD = new DirectoryEntry("WinNT://" +
                                    Environment.MachineName + ",computer");
                DirectoryEntry NewUser = AD.Children.Add(username, "user");
                NewUser.Invoke("SetPassword", new object[] { password });
                NewUser.Invoke("Put", new object[] { "Description", description });
                NewUser.CommitChanges();
                DirectoryEntry grp;

                grp = AD.Children.Find("Guests", "group");
                if (grp != null) { grp.Invoke("Add", new object[] { NewUser.Path.ToString() }); }
                Console.WriteLine("Account Created Successfully");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();

            }
        }

        /// <summary>
        /// Creates and starts a new IIS website using legacy IIS metabase methods.
        /// </summary>
        /// <param name="webserver">The web server name (usually "localhost").</param>
        /// <param name="serverComment">A description for the site.</param>
        /// <param name="serverBindings">Binding info in the format ":port:hostname".</param>
        /// <param name="homeDirectory">The physical path to the site's root.</param>
        /// <param name="appPool">The name of the application pool to assign.</param>
        /// <returns>1 on success.</returns>

        public static int StartWebsite(string webserver, string serverComment, string serverBindings, string homeDirectory, string appPool)
        {
            DirectoryEntry w3svc = new DirectoryEntry("IIS://localhost/w3svc");

            //Create a website object array
            object[] newsite = new object[] { serverComment, new object[] { serverBindings }, homeDirectory };



            //invoke IIsWebService.CreateNewSite
            object websiteId = (object)w3svc.Invoke("CreateNewSite", newsite);
            using (DirectoryEntry website = new DirectoryEntry(string.Format("IIS://{0}/w3svc/{1}", "localhost", websiteId)))
            {
                website.Invoke("Start", null);
            }
            return 1;
        }

      
        public static void removeSite(string siteName)
        {
            ServerManager serverMgr = new ServerManager();
            Site s1 = serverMgr.Sites[siteName]; // you can pass the site name or the site ID
            serverMgr.Sites.Remove(s1);
            serverMgr.CommitChanges();
        }


        public static void removeAppPool(string appName)
        {
            ServerManager serverMgr = new ServerManager();
            ApplicationPool a1 = serverMgr.ApplicationPools[appName];
            serverMgr.ApplicationPools.Remove(a1);
            serverMgr.CommitChanges();
        }
        public static void removeUser(string userNAme)
        {
            DirectoryEntry localDirectory = new DirectoryEntry("WinNT://" + Environment.MachineName.ToString());
            DirectoryEntries users = localDirectory.Children;
            DirectoryEntry user = users.Find(userNAme);
            users.Remove(user);

        }
        /// <summary>
        /// Creates a new website using the modern Microsoft.Web.Administration API and assigns an application pool.
        /// </summary>
        /// <param name="siteName">The site name.</param>
        /// <param name="bindingInformation">Binding string, e.g., "*:80:example.com".</param>
        /// <param name="physicalPath">The folder path for the website root.</param>
        /// <param name="appPool">The application pool to assign.</param>

        public static void CreateWebsite(string siteName, string bindingInformation, string physicalPath, string appPool)
        {
            try
            {
                using (ServerManager serverManager = new ServerManager())
                {
                    // Check if the site already exists
                    if (serverManager.Sites[siteName] != null)
                    {
                        Console.WriteLine($"A site with the name '{siteName}' already exists.");
                        return;
                    }

                    // Create a new site with HTTP binding
                    Site newSite = serverManager.Sites.Add(siteName, "http", bindingInformation, physicalPath);

                    // Assign the application pool
                    newSite.Applications["/"].ApplicationPoolName = appPool;

                    // Commit the changes to IIS
                    serverManager.CommitChanges();

                    System.Diagnostics.Debug.WriteLine($"Website '{siteName}' created successfully with HTTP binding.");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating website: {ex.Message}");
            }
        }

        public static void AddHostHeader(string hostHeader, string websiteID, int webPort = 80)
        {

            DirectoryEntry site = new DirectoryEntry("IIS://localhost/w3svc/" + websiteID);
            try
            {
                //Get everything currently in the serverbindings propery. 
                PropertyValueCollection serverBindings = site.Properties["ServerBindings"];

                //Add the new binding
                serverBindings.Add(":" + webPort.ToString() + ":" + hostHeader);

                //Create an object array and copy the content to this array
                Object[] newList = new Object[serverBindings.Count];
                serverBindings.CopyTo(newList, 0);

                //Write to metabase
                site.Properties["ServerBindings"].Value = newList;

                //Commit the changes
                site.CommitChanges();

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

        }



        public static void CreateAppPool(string metabasePath, string appPoolName, string version)
        {

            using (ServerManager serverManager = new ServerManager())
            {

                ApplicationPool newPool = serverManager.ApplicationPools.Add(appPoolName);
                newPool.ManagedRuntimeVersion = version;
                serverManager.CommitChanges();

            }

        }
        public static void StartAppPool(string PoolName)
        {
            ServerManager manager = new ServerManager();
            string status;
            //string DefaultSiteName = System.Web.Hosting.HostingEnvironment.ApplicationHost.GetSiteName();
            //Site defaultSite = manager.Sites[DefaultSiteName];
            //string appVirtaulPath = HttpRuntime.AppDomainAppVirtualPath;
            string mname = System.Environment.MachineName;
            string appPoolName = string.Empty;
            manager = ServerManager.OpenRemote(mname);
            ObjectState result = ObjectState.Unknown;

            ApplicationPoolCollection applicationPoolCollection = manager.ApplicationPools;

            foreach (ApplicationPool applicationPool in applicationPoolCollection)
            {
                //result = manager.ApplicationPools[appPoolName].State;
                result = applicationPool.State;
                if (PoolName == applicationPool.Name.ToString())
                {
                    if (result.ToString() == "Stopped")
                    {
                        applicationPool.Start();
                    }
                }
                //Console.ReadLine();
            }
        }
        public static void StopAppPool(string PoolName)
        {
            ServerManager manager = new ServerManager();
            string status;
            //string DefaultSiteName = System.Web.Hosting.HostingEnvironment.ApplicationHost.GetSiteName();
            //Site defaultSite = manager.Sites[DefaultSiteName];
            //string appVirtaulPath = HttpRuntime.AppDomainAppVirtualPath;
            string mname = System.Environment.MachineName;
            string appPoolName = string.Empty;
            manager = ServerManager.OpenRemote(mname);
            ObjectState result = ObjectState.Unknown;

            ApplicationPoolCollection applicationPoolCollection = manager.ApplicationPools;

            foreach (ApplicationPool applicationPool in applicationPoolCollection)
            {
                //result = manager.ApplicationPools[appPoolName].State;
                result = applicationPool.State;
                if (PoolName == applicationPool.Name.ToString())
                {
                    if (result.ToString() == "Started")
                    {
                        applicationPool.Stop();
                    }
                }
                //Console.ReadLine();
            }
        }
        public static string[] getAppPoolNameList()
        {
            ServerManager manager = new ServerManager();
            string status;
            //string DefaultSiteName = System.Web.Hosting.HostingEnvironment.ApplicationHost.GetSiteName();
            //Site defaultSite = manager.Sites[DefaultSiteName];
            //string appVirtaulPath = HttpRuntime.AppDomainAppVirtualPath;
            string mname = System.Environment.MachineName;
            string appPoolName = string.Empty;
            manager = ServerManager.OpenRemote(mname);
            ObjectState result = ObjectState.Unknown;
            string rVar = "";
            ApplicationPoolCollection applicationPoolCollection = manager.ApplicationPools;

            foreach (ApplicationPool applicationPool in applicationPoolCollection)
            {
                //result = manager.ApplicationPools[appPoolName].State;
                result = applicationPool.State;
                rVar = rVar + applicationPool.Name.ToString() + ",";

                //Console.ReadLine();
            }
            return rVar.Split(',');
        }

        public static string[] getWebsiteNameList()
        {
            ServerManager manager = new ServerManager();
            string status;
            //string DefaultSiteName = System.Web.Hosting.HostingEnvironment.ApplicationHost.GetSiteName();
            //Site defaultSite = manager.Sites[DefaultSiteName];
            //string appVirtaulPath = HttpRuntime.AppDomainAppVirtualPath;
            string mname = System.Environment.MachineName;
            string appPoolName = string.Empty;
            manager = ServerManager.OpenRemote(mname);
            ObjectState result = ObjectState.Unknown;
            string rVar = "";
            SiteCollection websiteCollection = manager.Sites;

            foreach (Site site in websiteCollection)
            {
                //result = applicationPool.State;
                rVar = rVar + site.Name.ToString() + ",";

            }
            return rVar.Split(',');

        }

        public static string getWebsiteListXML()
        {
            ServerManager manager = new ServerManager();
            string status;
            //string DefaultSiteName = System.Web.Hosting.HostingEnvironment.ApplicationHost.GetSiteName();
            //Site defaultSite = manager.Sites[DefaultSiteName];
            //string appVirtaulPath = HttpRuntime.AppDomainAppVirtualPath;
            string mname = System.Environment.MachineName;
            string appPoolName = string.Empty;
            manager = ServerManager.OpenRemote(mname);
            ObjectState result = ObjectState.Unknown;
            string rVar = "<newDataSet>";
            SiteCollection websiteCollection = manager.Sites;
            BindingCollection siteDomains;
            string bindingString = "";
            string stateString = "";
            string physicalPath = "";
            foreach (Site site in websiteCollection)
            {
                rVar = rVar + "<Table>";
                rVar = rVar + "<set>";
                rVar = rVar + "<name>" + site.Name.ToString() + "</name>";
                rVar = rVar + "<id>" + site.Id.ToString() + "</id>";


                try
                {
                    switch (site.State)
                    {
                        case ObjectState.Started:
                            rVar = rVar + "<state>Started</state>";
                            break;
                        case ObjectState.Stopped:
                            rVar = rVar + "<state>Stopped</state>";
                            break;
                        case ObjectState.Starting:
                            rVar = rVar + "<state>Starting</state>";
                            break;

                        case ObjectState.Stopping:
                            rVar = rVar + "<state>Stopping</state>";
                            break;

                        default:
                            rVar = rVar + "<state>Unknown</state>";
                            break;
                    }
                }
                catch
                {
                    rVar = rVar + "<state>Unknown</state>";
                }

                siteDomains = site.Bindings;
                foreach (Binding binding in siteDomains)
                {
                    bindingString = bindingString + binding.ToString() + ",";
                }
                rVar = rVar + "<bindings>" + bindingString + "</bindings>";
                bindingString = "";
                try
                {
                    physicalPath = site.Applications["/"].VirtualDirectories["/"].PhysicalPath;
                    rVar = rVar + "<physicalPath>" + physicalPath + "</physicalPath>";
                }
                catch
                {
                    rVar = rVar + "<physicalPath>Unknown</physicalPath>";
                }
                //rVar = rVar + "<path>" + site. + "</path>";
                rVar = rVar + "</set>";
                rVar = rVar + "</Table>";
            }
            return rVar + "</newDataSet>";

        }
        private static string iNull(string s)
        {
            try
            {
                return "";
            }
            catch
            {
                return "";
            }
        }
        public static string AppPoolStatus(string PoolName)
        {
            ServerManager manager = new ServerManager();
            string status;
            //string DefaultSiteName = System.Web.Hosting.HostingEnvironment.ApplicationHost.GetSiteName();
            //Site defaultSite = manager.Sites[DefaultSiteName];
            //string appVirtaulPath = HttpRuntime.AppDomainAppVirtualPath;
            string mname = System.Environment.MachineName;
            string appPoolName = string.Empty;
            manager = ServerManager.OpenRemote(mname);
            ObjectState result = ObjectState.Unknown;
            string rVar = "App Pool Name Not Found";
            ApplicationPoolCollection applicationPoolCollection = manager.ApplicationPools;

            foreach (ApplicationPool applicationPool in applicationPoolCollection)
            {
                //result = manager.ApplicationPools[appPoolName].State;
                result = applicationPool.State;
                if (PoolName == applicationPool.Name.ToString())
                {
                    rVar = result.ToString();

                }
                //Console.ReadLine();
            }
            return rVar;
        }
        public static void CreateVApp(string websiteId, string dirPath, string dirName, string appPool)
        {
            ServerManager mgr = new ServerManager();
            var app = mgr.Sites[websiteId.ToString()].Applications.Add(@"/" + dirName, dirPath);
            app.ApplicationPoolName = appPool;
            mgr.CommitChanges();
        }

        public static string GetVirtualDirectoriesXml(int websiteId)
        {
            using (ServerManager serverManager = new ServerManager())
            {
                // Find the site by ID
                var site = serverManager.Sites.FirstOrDefault(s => s.Id == websiteId);
                if (site == null)
                {
                    throw new ArgumentException($"Website with ID {websiteId} not found.");
                }

                // Create an XML document
                XmlDocument xmlDoc = new XmlDocument();
                XmlElement rootElement = xmlDoc.CreateElement("VirtualDirectories");
                xmlDoc.AppendChild(rootElement);

                // Loop through the site's applications
                foreach (var app in site.Applications)
                {
                    foreach (var vdir in app.VirtualDirectories)
                    {
                        XmlElement vdirElement = xmlDoc.CreateElement("VirtualDirectory");

                        XmlElement nameElement = xmlDoc.CreateElement("Name");
                        nameElement.InnerText = vdir.Path;
                        vdirElement.AppendChild(nameElement);

                        XmlElement pathElement = xmlDoc.CreateElement("PhysicalPath");
                        pathElement.InnerText = vdir.PhysicalPath;
                        vdirElement.AppendChild(pathElement);

                        rootElement.AppendChild(vdirElement);
                    }
                }

                // Convert XML document to string
                StringBuilder stringBuilder = new StringBuilder();
                using (XmlWriter writer = XmlWriter.Create(stringBuilder, new XmlWriterSettings { Indent = true }))
                {
                    xmlDoc.Save(writer);
                }

                return stringBuilder.ToString();
            }
        }

        public static string GetVirtualApplicationsXml(int websiteId)
        {
            using (ServerManager serverManager = new ServerManager())
            {
                // Find the site by ID
                var site = serverManager.Sites.FirstOrDefault(s => s.Id == websiteId);
                if (site == null)
                {
                    throw new ArgumentException($"Website with ID {websiteId} not found.");
                }

                // Create an XML document
                XmlDocument xmlDoc = new XmlDocument();
                XmlElement rootElement = xmlDoc.CreateElement("VirtualApplications");
                xmlDoc.AppendChild(rootElement);

                // Loop through the site's applications
                foreach (var app in site.Applications)
                {
                    XmlElement appElement = xmlDoc.CreateElement("Application");

                    XmlElement pathElement = xmlDoc.CreateElement("Path");
                    pathElement.InnerText = app.Path;
                    appElement.AppendChild(pathElement);

                    XmlElement physicalPathElement = xmlDoc.CreateElement("PhysicalPath");
                    physicalPathElement.InnerText = app.VirtualDirectories[0].PhysicalPath;
                    appElement.AppendChild(physicalPathElement);

                    rootElement.AppendChild(appElement);
                }

                // Convert XML document to string
                StringBuilder stringBuilder = new StringBuilder();
                using (XmlWriter writer = XmlWriter.Create(stringBuilder, new XmlWriterSettings { Indent = true }))
                {
                    xmlDoc.Save(writer);
                }

                return stringBuilder.ToString();
            }
        }
        public static void CreateVDir(string metabasePath, string vDirName, string physicalPath, string AppPoolId)
        {


            //  metabasePath is of the form "IIS://<servername>/<service>/<siteID>/Root[/<vdir>]"
            //    for example "IIS://localhost/W3SVC/1/Root" 
            //  vDirName is of the form "<name>", for example, "MyNewVDir"
            //  physicalPath is of the form "<drive>:\<path>", for example, "C:\Inetpub\Wwwroot"
            Console.WriteLine("\nCreating virtual directory {0}/{1}, mapping the Root application to {2}:",
                metabasePath, vDirName, physicalPath);

            try
            {
                DirectoryEntry site = new DirectoryEntry(metabasePath);
                string className = site.SchemaClassName.ToString();
                if ((className.EndsWith("Server")) || (className.EndsWith("VirtualDir")))
                {
                    DirectoryEntries vdirs = site.Children;
                    DirectoryEntry newVDir = vdirs.Add(vDirName, (className.Replace("Service", "Application")));
                    newVDir.Properties["ScriptMaps"][0] = @".htm,C:\Windows\Microsoft.NET\Framework\v4.0.30319\aspnet_isapi.dll,5,GET, HEAD, POST";
                    newVDir.Properties["AppPoolId"][0] = AppPoolId;
                    newVDir.Properties["Path"][0] = physicalPath;
                    newVDir.Properties["AccessScript"][0] = true;
                    // These properties are necessary for an application to be created.
                    newVDir.Properties["AppFriendlyName"][0] = vDirName;
                    newVDir.Properties["AppIsolated"][0] = "1";
                    newVDir.Properties["AppRoot"][0] = "/LM" + metabasePath.Substring(metabasePath.IndexOf("/", ("IIS://".Length)));

                    newVDir.CommitChanges();

                    Console.WriteLine(" Done.");
                }
                else
                    Console.WriteLine(" Failed. A virtual directory can only be created in a site or virtual directory node.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed in CreateVDir with the following exception: \n{0}", ex.Message);
            }
        }

        public static void createFTPDir(string ftpuser, string ftpDir)
        {
            try
            {
                DirectoryEntry objSite = new DirectoryEntry("IIS://Localhost/MSFTPSVC/1/Root");
                string strClass = objSite.SchemaClassName.ToString();
                DirectoryEntries iC = objSite.Children;

                DirectoryEntry makeSite = iC.Add(ftpuser, strClass.Replace("Service", "VirtualDir"));
                makeSite.Properties["Path"][0] = ftpDir;
                makeSite.Properties["AccessWrite"][0] = true;
                //makeSite.Properties["AuthAnonymous"][0] = false;

                makeSite.CommitChanges();
            }
            catch (Exception ex)
            {
                Console.WriteLine("FTP Create error: " + ex.Message + "<br>");
            }
        }

        public static void removeFTPDir(string ftpuser)
        {
            try
            {
                DirectoryEntry objSite = new DirectoryEntry("IIS://Localhost/MSFTPSVC/1/Root/" + ftpuser);
                string strClass = objSite.SchemaClassName.ToString();

                DirectoryEntries iC = objSite.Children;
                iC.Remove(objSite);

            }
            catch (Exception ex)
            {
                Console.WriteLine("FTP Create error: " + ex.Message + "<br>");
            }
        }

        public static void SetModifyWebPermissions(string dir, string user)
        {
            try
            {
                // Use icacls to grant permissions
                string arguments = $"\"{dir}\" /grant \"{user}:(OI)(CI)M\" /T";
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "icacls",
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process process = Process.Start(psi))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();

                    process.WaitForExit();

                    if (process.ExitCode == 0)
                    {
                        Console.WriteLine($"Permissions set successfully:\n{output}");
                    }
                    else
                    {
                        Console.WriteLine($"Error setting permissions:\n{error}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }


        public static void AssignVDirToAppPool(string metabasePath, string appPoolName)
        {

            //  metabasePath is of the form "IIS://<servername>/W3SVC/<siteID>/Root[/<vDir>]"
            //    for example "IIS://localhost/W3SVC/1/Root/MyVDir" 
            //  appPoolName is of the form "<name>", for example, "MyAppPool"
            Console.WriteLine("\nAssigning application {0} to the application pool named {1}:", metabasePath, appPoolName);

            try
            {
                DirectoryEntry vDir = new DirectoryEntry(metabasePath);
                string className = vDir.SchemaClassName.ToString();
                //if (className.EndsWith("VirtualDir"))
                //{
                object[] param = { 0, appPoolName, true };
                vDir.Invoke("AppCreate3", param);
                vDir.Properties["AppIsolated"][0] = "2";
                Console.WriteLine(" Done.");
                //}
                //else
                //Console.WriteLine(" Failed in AssignVDirToAppPool; only virtual directories can be assigned to application pools");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed in AssignVDirToAppPool with the following exception: \n{0}" + ex.Message);
            }
        }

        public static DataSet GetWebsitesInfo()
        {
            try
            {
                // Create a new DataSet to hold website information
                DataSet websitesDataSet = new DataSet("Websites");

                // Create a DataTable for storing website details
                DataTable websiteTable = new DataTable("Website");
                websiteTable.Columns.Add("Name", typeof(string));
                websiteTable.Columns.Add("ID", typeof(string));
                websiteTable.Columns.Add("State", typeof(string));
                websiteTable.Columns.Add("PhysicalPath", typeof(string));
                websiteTable.Columns.Add("Bindings", typeof(string));

                // Initialize ServerManager
                ServerManager manager = new ServerManager();
                SiteCollection websiteCollection = manager.Sites;

                // Loop through all websites
                foreach (Site site in websiteCollection)
                {
                    string name = site.Name;
                    string id = site.Id.ToString();
                    string state = site.State.ToString();
                    string physicalPath = string.Empty;
                    string bindings = string.Empty;

                    // Safely access site properties
                    try
                    {
                        physicalPath = site.Applications["/"].VirtualDirectories["/"].PhysicalPath ?? "N/A";
                        bindings = string.Join(", ", site.Bindings.Select(b =>
                            $"{b.Protocol}://{b.Host}:{b.EndPoint?.Port}"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error accessing site properties for {name}: {ex.Message}");
                    }

                    // Add data to the DataTable
                    websiteTable.Rows.Add(name, id, state, physicalPath, bindings);
                }

                // Add the DataTable to the DataSet
                websitesDataSet.Tables.Add(websiteTable);

                return websitesDataSet;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving websites: {ex.Message}");
                return null;
            }
        }



    }
}
