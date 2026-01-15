using Microsoft.Extensions.Configuration;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using NWTAOutlineAssist;
using System;
using System.Diagnostics;
using System.IO;


namespace NWTAOutlineAssistUI.Views
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class Generate_Outline : Page
    {
        string messageText;

        public Generate_Outline()
        {
            this.InitializeComponent();
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {

            using (var memStream = new MemoryStream())
            {
                using (TextWriter writer = new StreamWriter(memStream))
                {
                    bool success = false;
                    string errMsg = string.Empty;
                    var outlinePrint = new OutlinePrint(App.AppInstance.Configuration, writer);
                    try
                    {
                        outlinePrint.GenerateOutline();
                        writer.WriteLine("Outline created successfully!");
                        writer.Flush();
                        success = true;
                    }
                    catch (Exception ex)
                    {
                        errMsg = ex.Message;
                    }
                    memStream.Position = 0;
                    StreamReader reader = new StreamReader(memStream);
                    string text = reader.ReadToEnd();

                    messageText = text + "\n" + errMsg;
                    Bindings.Update();

                    if (success && OpenOutline.IsChecked == true)
                    {
                        var cfg = App.AppInstance.Configuration;
                        var path = cfg.FullPath(cfg.OutlineOutput);
                        try
                        {
                            Process.Start(new ProcessStartInfo { FileName = path, UseShellExecute = true });
                        }
                        catch (Exception ex)
                        {
                            App.AppInstance.MainWindow.ShowErrorDialog("Could not open document: " + path, ex);
                        }
                    }
                        
                }
            }
        }
    }
}
