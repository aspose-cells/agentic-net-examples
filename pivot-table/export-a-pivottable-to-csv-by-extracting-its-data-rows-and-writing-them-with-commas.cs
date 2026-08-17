// Title: C# – Export Aspose.Cells PivotTable DataRows to CSV
// Description: Demonstrates how to create a workbook, build a PivotTable that sums values, retrieve its DataBodyRange, convert the range to a string‑based DataTable, join each row with commas, and write the output to a CSV file while optionally saving the original Excel file.
// Keywords: Aspose.Cells C# PivotTable CSV export | Export PivotTable DataBodyRange | Aspose.Cells ExportDataTableAsString | PivotTable to CSV example | C# write CSV from pivot data | Aspose.Cells API CSV generation | pivot table data rows extraction | Aspose.Cells CSV file creation | C# Excel pivot export
// Common Searches: How to export a PivotTable created with Aspose.Cells to CSV in C# | Aspose.Cells get pivot table rows without headers for CSV | C# extract PivotTable DataBodyRange and save as comma‑separated file | Export Aspose.Cells pivot results to a text file | Aspose.Cells CSV export from pivot table data
// Developer Intent: Export the result rows of an Aspose.Cells PivotTable to a CSV file using C#.
// Use Cases: Generate a lightweight CSV report from pivot calculations for downstream data pipelines. | Provide a delimited‑text version of pivot results for systems that cannot consume Excel files. | Create automated CSV exports of summarized data while keeping the original workbook for visual review.
// AI Prompts: Write C# code with Aspose.Cells that extracts a PivotTable's DataBodyRange and saves it as a CSV file using commas as delimiters. | Explain the steps to convert a PivotTable range to a DataTable and then to CSV with Aspose.Cells. | Suggest performance tips for exporting large PivotTables to CSV with Aspose.Cells in C#.

using System;
using System.Data;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableToCsvExport
{
    // Demonstrates how to create a workbook, build a PivotTable that sums values, retrieve its DataBodyRange, convert the range to a string‑based DataTable, join each row with commas, and write the output to a CSV file while optionally saving the original Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Sample data for the pivot table
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Amount");
                dataSheet.Cells["A2"].PutValue("Food");
                dataSheet.Cells["B2"].PutValue(120);
                dataSheet.Cells["A3"].PutValue("Travel");
                dataSheet.Cells["B3"].PutValue(300);
                dataSheet.Cells["A4"].PutValue("Food");
                dataSheet.Cells["B4"].PutValue(80);
                dataSheet.Cells["A5"].PutValue("Travel");
                dataSheet.Cells["B5"].PutValue(150);

                // Add a worksheet for the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // Add the pivot table (source range, destination cell, name)
                int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:B5", "A1", "MyPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure the pivot table: rows = Category, data = Sum of Amount
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Layout in tabular form for easier export
                pivotTable.ShowInTabularForm();

                // Refresh and calculate the pivot data using the correct API
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Get the data body range of the pivot table (contains the result rows)
                CellArea dataBody = pivotTable.DataBodyRange;

                // Calculate row and column counts from CellArea
                int rowCount = dataBody.EndRow - dataBody.StartRow + 1;
                int columnCount = dataBody.EndColumn - dataBody.StartColumn + 1;

                // Export the data body range to a DataTable as strings (no column headers)
                DataTable dt = pivotSheet.Cells.ExportDataTableAsString(
                    dataBody.StartRow,
                    dataBody.StartColumn,
                    rowCount,
                    columnCount,
                    false);

                // Build CSV content
                StringBuilder csvBuilder = new StringBuilder();
                foreach (DataRow row in dt.Rows)
                {
                    // Join each column value with a comma
                    string line = string.Join(",", row.ItemArray);
                    csvBuilder.AppendLine(line);
                }

                // Write CSV to file
                string csvPath = "PivotExport.csv";
                File.WriteAllText(csvPath, csvBuilder.ToString());

                // Save the workbook (optional, to see the pivot table)
                string workbookPath = "PivotWithData.xlsx";
                workbook.Save(workbookPath);

                Console.WriteLine($"Pivot table data exported to CSV at: {Path.GetFullPath(csvPath)}");
                Console.WriteLine($"Workbook saved at: {Path.GetFullPath(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
