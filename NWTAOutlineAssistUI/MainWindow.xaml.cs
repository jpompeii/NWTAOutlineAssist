using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage.Pickers;
using Windows.Storage;
using Microsoft.Extensions.Configuration;
using Windows.Storage.AccessCache;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace NWTAOutlineAssist
{

    public enum ApplicationMode
    {
        NoCurrentOutline,
        CreatingOutline,
        EditingOutline
    }

    public sealed partial class MainWindow : Window
    {
        public ApplicationMode CurrentMode { get; private set; }
        public XamlRoot XmlRoot => this.rootWindow.XamlRoot;
        public MainWindow()
        {
            this.InitializeComponent();

            // NavigationViewControl.SelectedItem = NavigationViewControl.MenuItems.OfType<NavigationViewItem>().First();
            ContentFrame.Navigate(
                       typeof(Views.StartPage),
                       null,
                       new Microsoft.UI.Xaml.Media.Animation.EntranceNavigationTransitionInfo()
                       );

        }

        private void NavigationViewControl_ItemInvoked(NavigationView sender,
          NavigationViewItemInvokedEventArgs args)
        {
            if (args.IsSettingsInvoked == true)
            {
                ContentFrame.Navigate(typeof(Views.SettingsPage), null, args.RecommendedNavigationTransitionInfo);
            }
            else if (args.InvokedItemContainer != null && (args.InvokedItemContainer.Tag != null))
            {
                Type newPage = Type.GetType(args.InvokedItemContainer.Tag.ToString());
                ContentFrame.Navigate(
                       newPage,
                       null,
                       args.RecommendedNavigationTransitionInfo
                       );
            }
        }

        private void ContentFrame_Navigated(object sender, NavigationEventArgs e)
        {
            NavigationViewControl.IsBackEnabled = ContentFrame.CanGoBack;

            if (ContentFrame.SourcePageType == typeof(Views.SettingsPage))
            {
                // SettingsItem is not part of NavView.MenuItems, and doesn't have a Tag.
                NavigationViewControl.SelectedItem = (NavigationViewItem)NavigationViewControl.SettingsItem;
            }
            else if (ContentFrame.SourcePageType != null)
            {
                NavigationViewControl.SelectedItem = NavigationViewControl.MenuItems
                    .OfType<NavigationViewItem>()
                    .First(n => n.Tag.Equals(ContentFrame.SourcePageType.FullName.ToString()));
            }

            NavigationViewControl.Header = ((NavigationViewItem)NavigationViewControl.SelectedItem)?.Content?.ToString();
        }

        public void SetCurrentApplicationMode(ApplicationMode appMode)
        {
            string newPageTag = "NWTAOutlineAssist.Views.HomePage";
            CurrentMode = appMode;

            foreach (NavigationViewItem menuItem in NavigationViewControl.MenuItems)
            {
                bool visibility = false;
                // NavigationViewItem menuItem = (NavigationViewItem)item;
                if (menuItem.Tag.ToString() == "NWTAOutlineAssist.Views.StartPage")
                {
                    if (appMode == ApplicationMode.NoCurrentOutline)
                    {
                        visibility = true;
                        NavigationViewControl.SelectedItem = menuItem;
                        newPageTag = menuItem.Tag.ToString();
                    }
                }
                else if (menuItem.Tag.ToString() == "NWTAOutlineAssist.Views.NewOutline")
                {
                    if (appMode == ApplicationMode.CreatingOutline)
                    {
                        visibility = true;
                        NavigationViewControl.SelectedItem = menuItem;
                        newPageTag = menuItem.Tag.ToString();
                    }
                }
                else if (appMode == ApplicationMode.EditingOutline)
                {
                    visibility = true;
                }
                menuItem.Visibility = visibility ? Visibility.Visible : Visibility.Collapsed;
            }
            NavigationViewControl.IsSettingsVisible = appMode == ApplicationMode.EditingOutline;
            Type newPage = Type.GetType(newPageTag);
            if (newPage != null)
            {
                ContentFrame.Navigate(newPage, null, null);
            }
        }

        public async void OpenOutline()
        {
            FolderPicker openPicker = new();
            var window = App.AppInstance.MainWindow;
            var hWnd = WinRT.Interop.WindowNative.GetWindowHandle(window);
            WinRT.Interop.InitializeWithWindow.Initialize(openPicker, hWnd);

            // Set options for your folder picker
            openPicker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            openPicker.FileTypeFilter.Add(".yaml");
            openPicker.FileTypeFilter.Add(".xlxs");

            StorageFolder folder = await openPicker.PickSingleFolderAsync();
            if (folder != null)
            {
                var app = App.AppInstance;
                app.TestInputFile(folder.Path, "Outline.yaml");
                try
                {
                    app.OpenOutline(folder.Path, false);
                    SetCurrentApplicationMode(ApplicationMode.EditingOutline);
                }
                catch (Exception ex)
                {
                    ShowErrorDialog("An error occurred trying to open the outline.", ex);
                    return;
                }
            }
        }

        public async void ShowErrorDialog(Exception ex)
        {
            string baseMsg = ex.Message;
            if (ex.InnerException != null)
            {
                baseMsg += "\nCause: " + ex.InnerException.Message;
            }

            ContentDialog errorDialog = new ContentDialog
            {
                XamlRoot = this.XmlRoot,
                Title = "Application Error",
                Content = baseMsg,
                CloseButtonText = "OK"
            };

            ContentDialogResult result = await errorDialog.ShowAsync();
        }

        public async void ShowErrorDialog(string msg, Exception ex = null)
        {
            string baseMsg = msg;
            if (ex != null)
            {
                baseMsg += "\nCause: " + ex.Message;
                if (ex.InnerException != null)
                {
                    baseMsg += " Due to: " + ex.InnerException.Message;
                }
            }

            ContentDialog errorDialog = new ContentDialog
            {
                XamlRoot = this.XmlRoot,
                Title = "Application Error",
                Content = baseMsg,
                CloseButtonText = "OK"
            };

            ContentDialogResult result = await errorDialog.ShowAsync();
        }
    }
}
