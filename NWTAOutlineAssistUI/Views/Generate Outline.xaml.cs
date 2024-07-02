using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using NWTAOutlineAssist;
using System;
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
                    string errMsg = string.Empty;
                    var outlinePrint = new OutlinePrint(App.AppInstance.Configuration, writer);
                    try
                    {
                        outlinePrint.GenerateOutline();
                        writer.WriteLine("Outline created successfully!");
                        writer.Flush();
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
                }
            }
        }
    }
}
