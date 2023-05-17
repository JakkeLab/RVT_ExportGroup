using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Reflection;

namespace ExportSelectedAs
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Create a new Ribbon Tab
                string tabName = "Export As";
                application.CreateRibbonTab(tabName);
                // assembly
                Assembly assembly = Assembly.GetExecutingAssembly();
                // assembly path
                string assemblyPath = assembly.Location;

                // Create a new Ribbon Panel
                RibbonPanel ribbonPanel = application.CreateRibbonPanel(tabName, "Export Group");

                // Create a new Push Button
                PushButtonData buttonData = new PushButtonData("Export Groups as NWC", "Export\nGroup(.nwc)", assemblyPath, "ExportSelectedAs.Command");

                // Add the Push Button to the Ribbon Panel
                PushButton pushButton = ribbonPanel.AddItem(buttonData) as PushButton;

                // return result
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
