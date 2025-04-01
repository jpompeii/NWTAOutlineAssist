using Microsoft.Extensions.Configuration;
using Microsoft.UI.Xaml;
using Microsoft.Win32;
using System;
using System.IO;
using System.Security.AccessControl;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;


namespace NWTAOutlineAssist
{
    /// <summary>
    /// Provides application-specific behavior to supplement the default Application class.
    /// </summary>

    public partial class App : Application
    {
        public OAConfiguration Configuration { get; private set; }
        public string CurrentOutline { get; private set; }
        public static App AppInstance { get; private set; }
        public MainWindow MainWindow { get; private set; }
        
        

        /// <summary>
        /// Initializes the singleton application object.  This is the first line of authored code
        /// executed, and as such is the logical equivalent of main() or WinMain().
        /// </summary>
        public App()
        {
            Configuration = new OAConfiguration();
            this.InitializeComponent();
        }

        /// <summary>
        /// Invoked when the application is launched.
        /// </summary>
        /// <param name="args">Details about the launch request and process.</param>
        protected override void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
        {
            MainWindow = new MainWindow();
            AppInstance = this;

            var currOutline = ReadCurrentOutlineFromReg();
            ApplicationMode mode = ApplicationMode.NoCurrentOutline;
            if (currOutline is not null && File.Exists(currOutline + @"\Outline.yaml"))
            {
                try
                {
                    OpenOutline(currOutline, true);
                    mode = ApplicationMode.EditingOutline;
                }
                catch { }
            }

            MainWindow.Activate();
            MainWindow.SetCurrentApplicationMode(mode); 

            var appWin = MainWindow.AppWindow;
            appWin.Show();
            appWin.Resize(new Windows.Graphics.SizeInt32(1000, 700));
            appWin.Title = "NWTA Outline Assist";
        }

        public void EstablishNewOutline(OAConfiguration configuration)
        {
            var configFile = configuration.FullPath("Outline.yaml");
            using (var output = File.CreateText(configFile))
            {
                var serializer = new SerializerBuilder().Build();
                var yaml = serializer.Serialize(configuration);
                output.WriteLine(yaml);
            }
            WriteCurrentOutlineToReg(configuration.OutlineFolder);
            Configuration = configuration;
            MainWindow.SetCurrentApplicationMode(ApplicationMode.EditingOutline);
        }

        public void OpenOutline(string path, bool isCurrent)
        {
            var configFile = path + @"\Outline.yaml";
            var deserializer = new YamlDotNet.Serialization.DeserializerBuilder().Build();

            OAConfiguration cfg = deserializer.Deserialize<OAConfiguration>(File.ReadAllText(configFile));
            cfg.OutlineFolder = path;
            TestInputFile(path, cfg.RoleAssignments);
            TestInputFile(path, cfg.OutlineTemplate);
            TestInputFile(path, cfg.StaffRoster);
            var outlineOut = cfg.OutlineName.Trim() + " Outline.xlsx";
            if (File.Exists(cfg.FullPath(outlineOut)))
                cfg.OutlineOutput= outlineOut;

            Configuration = cfg;
            if (!isCurrent)
            { 
                WriteCurrentOutlineToReg(path); 
            }
        }

        public void TestInputFile(string path, string file)
        {
            if (!File.Exists(path + "\\" + file))
            {
                throw new ApplicationException($"Cannot open file: {file} in foler: {path}");
            }
        }

        public string ReadCurrentOutlineFromReg()
        {
            string result = null;
            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\NWTAOutlineAssist");
            //if it does exist, retrieve the stored values  
            if (key != null)
            {
                result = key.GetValue("CurrentOutline").ToString();
                key.Close();
            }
            return result;  
        }

        public void WriteCurrentOutlineToReg(string currentOutlineDir)
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\NWTAOutlineAssist", true);
            if (key == null)
            {
                string user = Environment.UserDomainName + "\\" + Environment.UserName;
                RegistrySecurity rs = new RegistrySecurity();

                // Allow the current user to read and delete the key.
                //
                rs.AddAccessRule(new RegistryAccessRule(user,
                    RegistryRights.ReadKey | RegistryRights.Delete | RegistryRights.SetValue | RegistryRights.CreateSubKey,
                    InheritanceFlags.None,
                    PropagationFlags.None,
                    AccessControlType.Allow)); key = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\NWTAOutlineAssist", RegistryKeyPermissionCheck.Default, rs);
            }
            key.SetValue("CurrentOutline", currentOutlineDir);
            key.Close();
        }

     
    }
}
