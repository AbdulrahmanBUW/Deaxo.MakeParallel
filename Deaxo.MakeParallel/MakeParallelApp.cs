using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using System.Drawing;

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
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                PushButtonData buttonData = new PushButtonData(
                    "MakeParallel",
                    "Make\nParallel",
                    assemblyPath,
                    "Deaxo.MakeParallel.Commands.MakeParallelCommand"
                )
                {
                    ToolTip = "Make two elements parallel by rotating the second element to match the first element's direction in XY plane",
                };

                // Load large and small icons using GDI → WPF conversion
                buttonData.LargeImage = LoadBitmapFromEmbeddedResource("Deaxo.MakeParallel.Resources.parallel32.png");
                buttonData.Image = LoadBitmapFromEmbeddedResource("Deaxo.MakeParallel.Resources.parallel16.png");

                // Add button to panel
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

        /// <summary>
        /// Loads an embedded PNG and converts it to BitmapSource for Revit ribbon.
        /// </summary>
        /// <param name="resourceName">Namespace + folder + filename</param>
        /// <returns>BitmapSource or null</returns>
        private BitmapSource LoadBitmapFromEmbeddedResource(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return null;

                using (var bmp = new Bitmap(stream))
                {
                    var bitmapSource = Imaging.CreateBitmapSourceFromHBitmap(
                        bmp.GetHbitmap(),
                        IntPtr.Zero,
                        System.Windows.Int32Rect.Empty,
                        BitmapSizeOptions.FromWidthAndHeight(bmp.Width, bmp.Height)
                    );
                    bitmapSource.Freeze();
                    return bitmapSource;
                }
            }
        }
    }
}