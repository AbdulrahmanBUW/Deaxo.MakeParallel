using Autodesk.Revit.UI;
using System;

namespace Deaxo.MakeParallel
{
    public class MakeParallelApp : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                string tabName = "DEAXO Tools";

                // Create Ribbon tab if it doesn't exist
                try { application.CreateRibbonTab(tabName); } catch { }

                // Create Ribbon panel
                RibbonPanel panel = application.CreateRibbonPanel(tabName, "Alignment Tools");

                // PushButton data
                string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                PushButtonData buttonData = new PushButtonData(
                    "MakeParallel",
                    "Make\nParallel",
                    assemblyPath,
                    "Deaxo.MakeParallel.Commands.MakeParallelCommand"
                )
                {
                    ToolTip = "Make two elements parallel by rotating the second element to match the first element's direction in XY plane",
                };

                // Add button to panel (without icons for now - you can add them later)
                panel.AddItem(buttonData);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("DEAXO - Make Parallel Ribbon Error", ex.Message);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}