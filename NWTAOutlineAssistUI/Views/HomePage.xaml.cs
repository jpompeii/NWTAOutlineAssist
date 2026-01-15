using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace NWTAOutlineAssist.Views
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class HomePage : Page
    {
        public OAConfiguration Configuration { get; set; }
        public string OutlineName { get { return "Outline: " + Configuration.OutlineName; } }
        private string errorText;
        public HomePage()
        {
            this.InitializeComponent();
            Configuration = App.AppInstance.Configuration;
        }
        private void OpenRoleAssignments_Click(object sender, RoutedEventArgs e)
        {
            var uri = Configuration.RoleAssignments.StartsWith("https") ? Configuration.RoleAssignments : Configuration.FullPath(Configuration.RoleAssignments);
            OpenDocument(uri);
        }

        private void OpenOutlineTemplate_Click(object sender, RoutedEventArgs e)
        {
            OpenDocument(Configuration.FullPath(Configuration.OutlineTemplate));
        }

        private void OpenStaffRoster_Click(object sender, RoutedEventArgs e)
        {
            OpenDocument(Configuration.FullPath(Configuration.StaffRoster));    
        }

        private void OpenOutline_Click(object sender, RoutedEventArgs e)
        {
            if (!String.IsNullOrEmpty(Configuration.OutlineOutput))
            {
                OpenDocument(Configuration.FullPath(Configuration.OutlineOutput));
            }
        }

        private void OpenDocument(string path)
        {
            if (!String.IsNullOrEmpty(path))
            {
                try
                {
                    Process.Start(new ProcessStartInfo { FileName = path, UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    errorText = "Could not open document: " + path;
                    App.AppInstance.MainWindow.ShowErrorDialog(errorText, ex);
                }
            }
        }   

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Configuration = new OAConfiguration();
            App.AppInstance.MainWindow.SetCurrentApplicationMode(ApplicationMode.NoCurrentOutline);
        }
    }
}
