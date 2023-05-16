using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using ExportSelectedAs.TabUI;
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

                //Create Show Button
                PushButtonData showButtonData = new PushButtonData("Show Window", "Show", assemblyPath, "ExportSelectedAs.Show");
                PushButton showButton = ribbonPanel.AddItem(showButtonData) as PushButton;

                //register dockablepane
                RegisterDockablePane(application);

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

        public Result RegisterDockablePane(UIControlledApplication application)
        {
            // dockablepaneviewer (customcontrol)
            MyDockablePane window = new MyDockablePane();

            // register in application with a new guid
            DockablePaneId dockID = new DockablePaneId(new Guid("{2ab6b447-1be8-4439-bbd1-c3e9d4037d64}"));
            try
            {
                application.RegisterDockablePane(dockID, "Dock Pane",
                    window as IDockablePaneProvider);

            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error Message", ex.Message);
                return Result.Failed;
            }
            return Result.Succeeded;
        }
    }

    public class CommandAvailability : IExternalCommandAvailability
    {
        public bool IsCommandAvailable(UIApplication app, CategorySet cate)
        {
            if (app.ActiveUIDocument == null)
            {
                //disable register btn
                return true;
            }
            //enable register btn
            return false;
        }
    }
    // external command class
    [Transaction(TransactionMode.Manual)]
    public class Show : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // dockable window id
                DockablePaneId id = new DockablePaneId(new Guid("{2ab6b447-1be8-4439-bbd1-c3e9d4037d64}"));
                DockablePane dockableWindow = commandData.Application.GetDockablePane(id);
                dockableWindow.Show();
            }
            catch (Exception ex)
            {
                // show error info dialog
                TaskDialog.Show("Info Message", ex.Message);
            }
            // return result
            return Result.Succeeded;
        }
    }
}
