using System;
using Autodesk.Navisworks.Api.Plugins;

namespace AutoNAVSearchSets
{
    [Plugin("AutoNAVSearchSets",
            "ACLP_VDC",
            ToolTip = "Automatic Search Set Creation for Navisworks",
            DisplayName = "AutoNav SearchSets")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class PluginMain : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            try
            {
                MainWindow mainWindow = new MainWindow();
                mainWindow.ShowDialog();
                return 0;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    "Error launching AutoNav SearchSets:\n\n" + ex.Message + "\n\n" + ex.StackTrace,
                    "AutoNav SearchSets Error",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
                return 1;
            }
        }
    }
}
