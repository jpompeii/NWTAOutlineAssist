using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using NWTAOutlineAssist;
using Windows.Storage.AccessCache;
using Windows.Storage.Pickers;
using Windows.Storage;
using System;
using System.IO;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using NWTAOutlineAssistUI;
using Microsoft.UI.Xaml.Data;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace NWTAOutlineAssist.Views
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class NewOutline : Page
    {
        public OAConfiguration Configuration { get; set; }
        private List<string> xlsxFiles = new();
        private List<string> templates = new();
        private string errorText;
        public NewOutline()
        {
            this.InitializeComponent();
            Configuration = new();
            Configuration.OutlineFolder = "Choose Folder...";
            var templatesDir = AppDomain.CurrentDomain.BaseDirectory + @"\Templates\Output";
            
            DirectoryInfo tempDir = new DirectoryInfo(templatesDir);
            foreach (var file in tempDir.EnumerateFiles())
            {
                if (file.Extension == ".xlsx")
                {
                    templates.Add(file.Name);
                }
            }
        }

        private async void Folder_Click(object sender, RoutedEventArgs e)
        {
            FolderPicker openPicker = new();
            var window = App.AppInstance.MainWindow;
            var hWnd = WinRT.Interop.WindowNative.GetWindowHandle(window);
            WinRT.Interop.InitializeWithWindow.Initialize(openPicker, hWnd);

            // Set options for your folder picker
            openPicker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            openPicker.FileTypeFilter.Add("*");

            StorageFolder folder = await openPicker.PickSingleFolderAsync();
            if (folder != null)
            {
                StorageApplicationPermissions.FutureAccessList.AddOrReplace("PickedFolderToken", folder);
                Configuration.OutlineFolder = folder.Path;
                
                xlsxFiles.Clear();
                DirectoryInfo selectedFolder = new DirectoryInfo(folder.Path);
                foreach (var file in selectedFolder.EnumerateFiles())
                {
                    if (file.Extension == ".xlsx" || file.Extension == ".csv")
                    {
                        xlsxFiles.Add(file.Name);
                    }
                }
                Bindings.Update();
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            App.AppInstance.MainWindow.SetCurrentApplicationMode(ApplicationMode.NoCurrentOutline);
        }

        private void Create_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OutlineCreator creator = new(Configuration);
                creator.CreateOutline();
                App.AppInstance.EstablishNewOutline(Configuration);
            }
            catch (Exception ex)
            {
                // errorText = ex.Message;
                App.AppInstance.MainWindow.ShowErrorDialog("Could not create outline.", ex);
                Bindings.Update();
            }
        }
    }
}
