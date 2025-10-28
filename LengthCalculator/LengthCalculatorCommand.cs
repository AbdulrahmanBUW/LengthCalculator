using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

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
}