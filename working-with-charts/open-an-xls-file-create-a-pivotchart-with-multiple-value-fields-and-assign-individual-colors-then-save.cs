// Title: C# example: Open an XLS workbook, add a PivotTable with Sales and Quantity fields, create a linked PivotChart and set custom series colors using Aspose.Cells
// Common Searches: Aspose.Cells C# create pivot table with multiple data fields and generate a linked pivot chart | how to set individual series colors in a pivot chart using Aspose.Cells for .NET | convert XLS to XLSX while adding a colored pivot chart in C# | C# Aspose.Cells example for adding a column chart linked to a pivot table | assign custom foreground colors to NSeries in an Aspose.Cells chart
// Tags: Aspose.Cells create pivot table C# | Aspose.Cells add pivot chart linked to table | Aspose.Cells set series color NSeries | Aspose.Cells convert XLS to XLSX | Aspose.Cells column chart multiple value fields

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace PivotChartExample
{
    // The sample loads or creates an XLS workbook, builds a PivotTable with Category rows, SubCategory columns, and Sales and Quantity as data fields, refreshes the pivot, adds a column PivotChart linked to the table, assigns CornflowerBlue to the Sales series and OrangeRed to the Quantity series, sets a chart title, and saves the result as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust as needed
            string inputPath = "input.xls";
            string outputPath = "PivotChartResult.xlsx";

            try
            {
                // Ensure the input file exists; if not, create a simple sample workbook
                if (!File.Exists(inputPath))
                {
                    var sampleWb = new Workbook();
                    var sampleSheet = sampleWb.Worksheets[0];

                    // Sample data: Category, Sub‑Category, Sales, Quantity
                    sampleSheet.Cells["A1"].PutValue("Category");
                    sampleSheet.Cells["B1"].PutValue("SubCategory");
                    sampleSheet.Cells["C1"].PutValue("Sales");
                    sampleSheet.Cells["D1"].PutValue("Quantity");

                    sampleSheet.Cells["A2"].PutValue("Food");
                    sampleSheet.Cells["B2"].PutValue("Fruit");
                    sampleSheet.Cells["C2"].PutValue(1200);
                    sampleSheet.Cells["D2"].PutValue(30);

                    sampleSheet.Cells["A3"].PutValue("Food");
                    sampleSheet.Cells["B3"].PutValue("Vegetable");
                    sampleSheet.Cells["C3"].PutValue(800);
                    sampleSheet.Cells["D3"].PutValue(20);

                    sampleSheet.Cells["A4"].PutValue("Beverage");
                    sampleSheet.Cells["B4"].PutValue("Tea");
                    sampleSheet.Cells["C4"].PutValue(500);
                    sampleSheet.Cells["D4"].PutValue(15);

                    // Save as Excel 97‑2003 format
                    sampleWb.Save(inputPath, SaveFormat.Excel97To2003);
                }

                // Load the existing XLS workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0]; // assume first sheet contains source data

                // -------------------------------------------------
                // 1. Create a PivotTable (if the workbook does not already contain one)
                // -------------------------------------------------
                int lastRow = sheet.Cells.MaxDataRow;
                int lastCol = sheet.Cells.MaxDataColumn;
                string sourceRange = $"={sheet.Name}!A1:{CellsHelper.ColumnIndexToName(lastCol)}{lastRow + 1}";
                string destCell = "E3";

                // Use the classic overload (still valid)
                int pivotIndex = sheet.PivotTables.Add(sourceRange, destCell, "MyPivotTable");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // -------------------------------------------------
                // 2. Configure the PivotTable – multiple value fields
                // -------------------------------------------------
                // Row field (Category)
                pivot.AddFieldToArea(PivotFieldType.Row, 0); // column index 0

                // Column field (SubCategory)
                pivot.AddFieldToArea(PivotFieldType.Column, 1); // column index 1

                // Add two value fields (Sales and Quantity)
                pivot.AddFieldToArea(PivotFieldType.Data, 2);
                pivot.AddFieldToArea(PivotFieldType.Data, 3);

                // Set custom captions for the data fields (optional)
                pivot.DataFields[0].Name = "Sales";
                pivot.DataFields[1].Name = "Quantity";

                // Refresh and calculate the pivot data
                pivot.RefreshData();
                pivot.CalculateData();

                // -------------------------------------------------
                // 3. Create a PivotChart linked to the PivotTable
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Link the chart to the pivot table
                chart.PivotSource = $"{sheet.Name}!{pivot.Name}";

                // Refresh the chart so that series are generated from the pivot table
                chart.RefreshPivotData();

                // -------------------------------------------------
                // 4. Assign individual colors to each value field (series)
                // -------------------------------------------------
                if (chart.NSeries.Count >= 2)
                {
                    // First value field – Sales
                    chart.NSeries[0].Area.ForegroundColor = Color.CornflowerBlue;
                    chart.NSeries[0].Area.BackgroundColor = Color.Transparent;

                    // Second value field – Quantity
                    chart.NSeries[1].Area.ForegroundColor = Color.OrangeRed;
                    chart.NSeries[1].Area.BackgroundColor = Color.Transparent;
                }
                else
                {
                    Console.WriteLine($"Warning: Expected at least two data series in the chart, but found {chart.NSeries.Count}");
                }

                // Optional: set chart title
                chart.Title.Text = "Sales vs Quantity Pivot Chart";

                // -------------------------------------------------
                // 5. Save the workbook
                // -------------------------------------------------
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
