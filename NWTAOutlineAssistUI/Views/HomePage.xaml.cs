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
        private string errorText;
        public HomePage()
        {
            this.InitializeComponent();
            Configuration = App.AppInstance.Configuration;
        }
        private void OpenRoleAssignments_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(new ProcessStartInfo { FileName = Configuration.FullPath(Configuration.RoleAssignments), UseShellExecute = true });
        }

        private void OpenOutlineTemplate_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(new ProcessStartInfo { FileName = Configuration.FullPath(Configuration.OutlineTemplate), UseShellExecute = true });
        }

        private void OpenStaffRoster_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(new ProcessStartInfo { FileName = Configuration.FullPath(Configuration.StaffRoster), UseShellExecute = true });
        }

        private void OpenOutline_Click(object sender, RoutedEventArgs e)
        {
            if (!String.IsNullOrEmpty(Configuration.OutlineOutput))
            {
                Process.Start(new ProcessStartInfo { FileName = Configuration.FullPath(Configuration.OutlineOutput), UseShellExecute = true });
            }
        }
    }
}
