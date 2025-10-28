using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using Autodesk.Revit.DB;

namespace LengthCalculator
{
    public partial class LengthCalculatorWindow : Window
    {
        private readonly Document _doc;
        private readonly List<Element> _elements;
        private List<ElementLengthData> _results;

        public LengthCalculatorWindow(Document doc, List<Element> elements)
        {
            InitializeComponent();
            _doc = doc;
            _elements = elements;

            // Auto-calculate if elements are already selected
            if (_elements != null && _elements.Count > 0)
            {
                CalculateLengths();
            }
        }

        private void BtnCalculate_Click(object sender, RoutedEventArgs e)
        {
            CalculateLengths();
        }

        private void CalculateLengths()
        {
            try
            {
                if (_elements == null || _elements.Count == 0)
                {
                    MessageBox.Show(
                        "No elements selected. Please select elements with a Length parameter first.",
                        "No Selection",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                _results = new List<ElementLengthData>();
                double totalInternalUnits = 0.0;
                int countedElements = 0;

                foreach (var element in _elements)
                {
                    var data = new ElementLengthData
                    {
                        ElementName = GetElementName(element),
                        Size = GetSizeParameter(element)
                    };

                    // Find length parameter
                    var lengthInfo = FindLengthParameter(element);
                    data.LengthValue = lengthInfo.Value;
                    data.ParameterName = lengthInfo.ParameterName;
                    data.Source = lengthInfo.Source;

                    if (data.LengthValue.HasValue)
                    {
                        totalInternalUnits += data.LengthValue.Value;
                        countedElements++;

                        var (convertedValue, unitSymbol) = FormatInProjectUnits(data.LengthValue.Value, _doc);
                        data.LengthDisplay = $"{convertedValue:F4} {unitSymbol}";
                    }
                    else
                    {
                        data.LengthDisplay = "NO LENGTH PARAM";
                    }

                    _results.Add(data);
                }

                // Update summary statistics
                txtTotalElements.Text = _elements.Count.ToString();
                txtWithLength.Text = countedElements.ToString();

                var (totalConverted, totalUnit) = FormatInProjectUnits(totalInternalUnits, _doc);
                txtTotalLengthLabel.Text = $"TOTAL LENGTH ({totalUnit.ToUpper()})";
                txtTotalLength.Text = $"{totalConverted:F4}";

                // Calculate alternative units
                double totalMeters = totalInternalUnits * 0.3048;
                double totalFeet = totalInternalUnits;
                txtAlternativeUnits.Text = $"{totalMeters:F2} m / {totalFeet:F2} ft";

                // Bind to DataGrid
                dgResults.ItemsSource = _results;

                // Enable export buttons
                btnCopy.IsEnabled = true;
                btnExport.IsEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error calculating lengths: {ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private string GetElementName(Element element)
        {
            try
            {
                if (element is FamilyInstance familyInstance && familyInstance.Symbol != null)
                {
                    return $"{familyInstance.Symbol.FamilyName} : {element.Name}";
                }

                return element.Name ?? "Unknown";
            }
            catch
            {
                return "Unknown";
            }
        }

        private string GetSizeParameter(Element element)
        {
            try
            {
                var sizeParam = element.LookupParameter("Size");
                if (sizeParam != null)
                {
                    return sizeParam.AsString() ?? string.Empty;
                }
            }
            catch { }

            return string.Empty;
        }

        private (double? Value, string ParameterName, string Source) FindLengthParameter(Element element)
        {
            try
            {
                // Try instance "Length" parameter first
                var lengthParam = element.LookupParameter("Length");
                double? value = TryReadParameterValue(lengthParam);
                if (value.HasValue)
                {
                    return (value, lengthParam?.Definition?.Name ?? "Length", "Instance");
                }

                // Search instance parameters for anything containing "length"
                foreach (Parameter param in element.Parameters)
                {
                    string paramName = param?.Definition?.Name;
                    if (!string.IsNullOrEmpty(paramName) && paramName.ToLower().Contains("length"))
                    {
                        value = TryReadParameterValue(param);
                        if (value.HasValue)
                        {
                            return (value, paramName, "Instance");
                        }
                    }
                }

                // Try type parameters
                if (_doc.GetElement(element.GetTypeId()) is ElementType elementType)
                {
                    lengthParam = elementType.LookupParameter("Length");
                    value = TryReadParameterValue(lengthParam);
                    if (value.HasValue)
                    {
                        return (value, lengthParam?.Definition?.Name ?? "Length", "Type");
                    }

                    foreach (Parameter param in elementType.Parameters)
                    {
                        string paramName = param?.Definition?.Name;
                        if (!string.IsNullOrEmpty(paramName) && paramName.ToLower().Contains("length"))
                        {
                            value = TryReadParameterValue(param);
                            if (value.HasValue)
                            {
                                return (value, paramName, "Type");
                            }
                        }
                    }
                }
            }
            catch { }

            return (null, null, null);
        }

        private double? TryReadParameterValue(Parameter param)
        {
            try
            {
                if (param == null || !param.HasValue)
                    return null;

                var storageType = param.StorageType;

                switch (storageType)
                {
                    case StorageType.Double:
                        return param.AsDouble();

                    case StorageType.Integer:
                        return Convert.ToDouble(param.AsInteger());

                    case StorageType.String:
                        string str = param.AsString();
                        if (string.IsNullOrWhiteSpace(str))
                            return null;

                        // Try to extract numeric value from string
                        var match = Regex.Match(str, @"[-+]?\d+[,\.\d]*");
                        if (match.Success)
                        {
                            string numStr = match.Value.Replace(',', '.');
                            if (double.TryParse(numStr, out double result))
                            {
                                return result;
                            }
                        }
                        break;
                }
            }
            catch { }

            return null;
        }

        private (double ConvertedValue, string UnitSymbol) FormatInProjectUnits(double valueFeet, Document doc)
        {
            try
            {
                var units = doc.GetUnits();

                // Revit 2021+
                var formatOptions = units.GetFormatOptions(SpecTypeId.Length);
                var unitTypeId = formatOptions.GetUnitTypeId();
                double convertedValue = UnitUtils.ConvertFromInternalUnits(valueFeet, unitTypeId);

                string symbol = "ft";
                if (unitTypeId == UnitTypeId.Millimeters)
                    symbol = "mm";
                else if (unitTypeId == UnitTypeId.Centimeters)
                    symbol = "cm";
                else if (unitTypeId == UnitTypeId.Meters)
                    symbol = "m";
                else if (unitTypeId == UnitTypeId.Feet)
                    symbol = "ft";
                else if (unitTypeId == UnitTypeId.Inches)
                    symbol = "in";

                return (convertedValue, symbol);
            }
            catch
            {
                return (valueFeet, "ft");
            }
        }

        private void BtnCopy_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sb = new StringBuilder();

                // Add summary
                sb.AppendLine($"Total elements selected: {txtTotalElements.Text}");
                sb.AppendLine($"Elements with length param found: {txtWithLength.Text}");
                sb.AppendLine($"Total length: {txtTotalLength.Text} | {txtAlternativeUnits.Text}");
                sb.AppendLine();
                sb.AppendLine("Per-element:");
                sb.AppendLine();

                // Add element details
                foreach (var item in _results)
                {
                    sb.AppendLine($"{item.ElementName} (Size: {item.Size}) : {item.LengthDisplay}");
                }

                Clipboard.SetText(sb.ToString());

                MessageBox.Show(
                    "Results copied to clipboard!",
                    "Copied",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Could not copy to clipboard: {ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_results == null || _results.Count == 0)
                {
                    MessageBox.Show(
                        "No data to export. Please calculate first.",
                        "Export Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                // Create Excel application
                var excelApp = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible = true
                };

                var workbook = excelApp.Workbooks.Add();
                var worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];

                // Headers
                var (_, unitSymbol) = FormatInProjectUnits(1.0, _doc);
                worksheet.Cells[1, 1] = "Element Name";
                worksheet.Cells[1, 2] = "Size";
                worksheet.Cells[1, 3] = $"Length ({unitSymbol})";
                worksheet.Cells[1, 4] = "Parameter Name";
                worksheet.Cells[1, 5] = "Source";

                // Make headers bold
                var headerRange = worksheet.Range["A1", "E1"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                // Summary row
                double totalInternalUnits = _results.Where(r => r.LengthValue.HasValue).Sum(r => r.LengthValue.Value);
                double totalMm = totalInternalUnits * 0.3048 * 1000;
                double totalM = totalInternalUnits * 0.3048;
                double totalFt = totalInternalUnits;

                worksheet.Cells[2, 1] = "TOTALS";
                worksheet.Cells[2, 4] = $"{totalMm:F2} mm";
                worksheet.Cells[2, 5] = $"{totalM:F4} m";
                worksheet.Range["A2", "E2"].Font.Bold = true;

                // Data rows
                int row = 4;
                foreach (var item in _results)
                {
                    worksheet.Cells[row, 1] = item.ElementName;
                    worksheet.Cells[row, 2] = item.Size;
                    worksheet.Cells[row, 3] = item.LengthDisplay;
                    worksheet.Cells[row, 4] = item.ParameterName ?? "N/A";
                    worksheet.Cells[row, 5] = item.Source ?? "N/A";
                    row++;
                }

                // Auto-fit columns
                worksheet.Columns.AutoFit();

                MessageBox.Show(
                    "Data exported to Excel successfully!",
                    "Export Complete",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Failed to export to Excel: {ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }

    public class ElementLengthData
    {
        public string ElementName { get; set; }
        public string Size { get; set; }
        public double? LengthValue { get; set; }
        public string LengthDisplay { get; set; }
        public string ParameterName { get; set; }
        public string Source { get; set; }
    }
}