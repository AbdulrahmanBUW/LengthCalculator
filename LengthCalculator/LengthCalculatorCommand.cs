using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace LengthCalculator
{
    [Transaction(TransactionMode.ReadOnly)]
    [Regeneration(RegenerationOption.Manual)]
    public class LengthCalculatorCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // Get selected elements
                var selectedIds = uidoc.Selection.GetElementIds();
                var selectedElements = selectedIds
                    .Select(id => doc.GetElement(id))
                    .Where(e => e != null)
                    .ToList();

                // Show the main window
                var window = new LengthCalculatorWindow(doc, selectedElements);
                window.ShowDialog();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    /// <summary>
    /// Application class to add the button to Revit ribbon with embedded icon
    /// </summary>
    public class LengthCalculatorApp : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Create a ribbon panel
                string tabName = "Add-Ins";
                RibbonPanel ribbonPanel = application.CreateRibbonPanel(tabName, "Length Tools");

                // Get the assembly path
                string assemblyPath = Assembly.GetExecutingAssembly().Location;

                // Create push button data
                PushButtonData buttonData = new PushButtonData(
                    "LengthCalculator",
                    "Length\nCalculator",
                    assemblyPath,
                    "LengthCalculator.LengthCalculatorCommand");

                // Set tooltip
                buttonData.ToolTip = "Calculate total length of selected elements";
                buttonData.LongDescription = "Select pipes, ducts, beams, or other elements with length parameters and calculate the total length with detailed breakdown.";

                // Load embedded icon
                BitmapImage icon = LoadEmbeddedImage("LengthCalculator.length.png");
                if (icon != null)
                {
                    buttonData.LargeImage = icon;
                    buttonData.Image = icon;
                }

                // Add button to panel
                PushButton pushButton = ribbonPanel.AddItem(buttonData) as PushButton;

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to load Length Calculator: {ex.Message}");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        /// <summary>
        /// Load an image from embedded resources
        /// </summary>
        /// <param name="resourceName">Full resource name including namespace (e.g., "LengthCalculator.adjust.png")</param>
        /// <returns>BitmapImage or null if not found</returns>
        private BitmapImage LoadEmbeddedImage(string resourceName)
        {
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();

                // Get the embedded resource stream
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        // Try to list available resources for debugging
                        var availableResources = assembly.GetManifestResourceNames();
                        TaskDialog.Show("Debug", $"Resource '{resourceName}' not found.\n\nAvailable resources:\n{string.Join("\n", availableResources)}");
                        return null;
                    }

                    BitmapImage bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.StreamSource = stream;
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    bitmap.Freeze();

                    return bitmap;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error Loading Icon", $"Failed to load embedded icon: {ex.Message}");
                return null;
            }
        }
    }
}