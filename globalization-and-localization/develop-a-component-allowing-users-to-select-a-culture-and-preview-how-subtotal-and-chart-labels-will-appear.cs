// Title: Create a culture‑aware Excel workbook with subtotal row and chart label preview using Aspose.Cells for .NET
// Common Searches: how to generate an Excel subtotal with locale‑specific currency formatting using Aspose.Cells in C# | Aspose.Cells example to preview chart axis labels for a selected culture in a .NET application | set thread culture before exporting an Excel file with Aspose.Cells for culture‑aware number formats | display culture‑specific currency symbols in Excel chart value axis with Aspose.Cells C# | preview Excel workbook with different cultures without opening Excel using Aspose.Cells
// Tags: culture-specific number format Aspose.Cells | subtotal row currency formatting C# | chart axis locale formatting Aspose.Cells | thread culture Excel export .NET | preview workbook for selected culture Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample sets the thread's culture, creates a workbook, fills sample data, applies a culture‑specific currency number format to the amount column, adds a subtotal formula, recalculates formulas, inserts a column chart with category and value data, optionally formats the chart's value axis for the selected locale, ensures the output directory exists, and saves the workbook, demonstrating how to preview subtotal and chart labels for any chosen culture using Aspose.Cells for .NET.
class CulturePreview
{
    // Generates a workbook preview using the specified culture.
    public static void PreviewCulture(string cultureName, string outputPath)
    {
        try
        {
            // Set thread culture for number and date formatting.
            CultureInfo culture = new CultureInfo(cultureName);
            Thread.CurrentThread.CurrentCulture = culture;
            Thread.CurrentThread.CurrentUICulture = culture;

            // Create a new workbook and get the first worksheet.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Data";

            // Populate sample data.
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Amount");
            ws.Cells["A2"].PutValue("Food");
            ws.Cells["B2"].PutValue(1234.56);
            ws.Cells["A3"].PutValue("Food");
            ws.Cells["B3"].PutValue(789.01);
            ws.Cells["A4"].PutValue("Transport");
            ws.Cells["B4"].PutValue(456.78);
            ws.Cells["A5"].PutValue("Transport");
            ws.Cells["B5"].PutValue(321.09);

            // Apply culture‑specific number format to the amount column.
            Style amountStyle = wb.CreateStyle();
            amountStyle.Custom = culture.NumberFormat.CurrencySymbol + "#,##0.00";
            StyleFlag styleFlag = new StyleFlag { NumberFormat = true };
            ws.Cells["B2:B5"].SetStyle(amountStyle, styleFlag);

            // Insert a subtotal row (simple sum of all amounts).
            ws.Cells["A6"].PutValue("Subtotal");
            ws.Cells["B6"].Formula = "SUM(B2:B5)";

            // Recalculate formulas to reflect the current culture.
            wb.CalculateFormula();

            // Add a column chart to visualize the data.
            int chartIdx = ws.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];
            chart.Title.Text = "Amount by Category";

            // Set chart data source.
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // (Optional) Apply culture‑specific number format to the chart's value axis.
            // Axis valueAxis = chart.ValueAxis;
            // if (valueAxis != null)
            // {
            //     valueAxis.NumberFormat = culture.NumberFormat.CurrencySymbol + "#,##0.00";
            // }

            // Ensure the output directory exists.
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook for preview.
            wb.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error generating preview: {ex.Message}");
        }
    }

    static void Main(string[] args)
    {
        try
        {
            // Example: user selects French (France) culture.
            string selectedCulture = "fr-FR";
            string outputFile = "CulturePreview.xlsx";

            // Ensure the template or output path is valid before proceeding.
            PreviewCulture(selectedCulture, outputFile);
            Console.WriteLine($"Workbook saved to '{outputFile}' using culture '{selectedCulture}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
