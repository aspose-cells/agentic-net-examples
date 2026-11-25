// Title: C# Aspose.Cells example: link chart data label number format to source cells for dynamic inheritance
// Common Searches: c# aspose.cells link chart data label format to source cell style | how to make Excel chart data labels inherit number format from cells using Aspose.Cells | Aspose.Cells DataLabels NumberFormatLinked property example in C# | update cell number format and automatically refresh chart data labels with Aspose.Cells | dynamic chart label formatting based on cell style Aspose.Cells .NET
// Tags: Aspose.Cells chart label format inheritance | NumberFormatLinked API C# | custom number format for Excel cells Aspose.Cells | auto update chart labels after cell style modification | column chart with linked label formatting using Aspose.Cells | dynamic cell style propagation to chart data labels

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelNumberFormatLink
{
    // The sample creates a workbook with sales data, applies a custom number format to the value cells, builds a column chart, enables data labels, and sets DataLabels.NumberFormatLinked to true so the labels inherit the cells' formatting. After saving, the cell format is changed, the workbook is recalculated and saved again, demonstrating that the chart data labels automatically reflect the updated number format. A verification step reads back the file to confirm the linked formatting behavior.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and populate it with sample data.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Header
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");

                // Data rows
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");

                sheet.Cells["B2"].PutValue(1234.56);
                sheet.Cells["B3"].PutValue(7890.12);
                sheet.Cells["B4"].PutValue(3456.78);

                // -----------------------------------------------------------------
                // 2. Apply an initial number format to the source cells (B2:B4).
                // -----------------------------------------------------------------
                Style sourceStyle = workbook.CreateStyle();
                sourceStyle.Custom = "#,##0.00"; // two decimal places

                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true; // only change number format

                // Apply the style to the range B2:B4
                Aspose.Cells.Range sourceRange = sheet.Cells.CreateRange("B2:B4");
                sourceRange.ApplyStyle(sourceStyle, flag);

                // -----------------------------------------------------------------
                // 3. Create a column chart that uses the data.
                // -----------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Monthly Sales";

                // Set the series data and category (X‑axis) data.
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // -----------------------------------------------------------------
                // 4. Enable data labels and link their number format to the source cells.
                // -----------------------------------------------------------------
                Series series = chart.NSeries[0];
                // DataLabels are automatically created when accessed.
                DataLabels dataLabels = series.DataLabels;
                dataLabels.ShowValue = true;               // show the numeric value
                dataLabels.NumberFormatLinked = true;      // link number format to source cells

                // -----------------------------------------------------------------
                // 5. Save the workbook – data labels should reflect the initial format.
                // -----------------------------------------------------------------
                string initialPath = "ChartDataLabelLinked_initial.xlsx";
                workbook.Save(initialPath);
                Console.WriteLine($"Workbook saved with initial number format: {initialPath}");

                // -----------------------------------------------------------------
                // 6. Change the number format of the source cells (B2:B4) to a different pattern.
                // -----------------------------------------------------------------
                Style updatedStyle = workbook.CreateStyle();
                updatedStyle.Custom = "#,##0.0000"; // four decimal places

                // Re‑apply the new style to the same range.
                sourceRange.ApplyStyle(updatedStyle, flag);

                // -----------------------------------------------------------------
                // 7. Re‑calculate (not strictly required for static values) and save again.
                // -----------------------------------------------------------------
                workbook.CalculateFormula(); // ensures any formulas are up‑to‑date
                string updatedPath = "ChartDataLabelLinked_updated.xlsx";
                workbook.Save(updatedPath);
                Console.WriteLine($"Workbook saved after updating number format: {updatedPath}");

                // -----------------------------------------------------------------
                // 8. Verify that the data label format has changed by reading back.
                // -----------------------------------------------------------------
                if (!File.Exists(updatedPath))
                {
                    Console.WriteLine($"Error: File \"{updatedPath}\" was not found.");
                    return;
                }

                try
                {
                    Workbook verifyWb = new Workbook(updatedPath);
                    Worksheet verifySheet = verifyWb.Worksheets[0];
                    Chart verifyChart = verifySheet.Charts[0];
                    string labelFormat = verifyChart.NSeries[0].DataLabels.NumberFormat;
                    Console.WriteLine($"Data label number format after update (should be empty because it's linked): \"{labelFormat}\"");
                    Console.WriteLine("If NumberFormatLinked is true, the label inherits the cell format dynamically.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Verification failed: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
