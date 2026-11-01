// Title: Disable automatic calculation, import a DataTable into an Excel worksheet, and manually recalculate formulas with Aspose.Cells for .NET
// Common Searches: how to turn off auto calculation in Aspose.Cells C# workbook | import DataTable into Excel using Aspose.Cells and keep formulas unchanged | manually invoke workbook.CalculateFormula after loading data in Aspose.Cells | disable formula auto‑recalculation before populating worksheet with database data | Aspose.Cells C# example for loading database rows and triggering calculation later
// Tags: disable auto calculation Aspose.Cells | DataTable import Aspose.Cells worksheet | manual formula calculation Aspose.Cells | populate Excel from database Aspose.Cells | calculate formulas after data import Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, disables automatic formula calculation, imports a DataTable (simulating database data) into the first worksheet, manually triggers workbook.CalculateFormula to recalculate all formulas, and saves the file as Result.xlsx with error handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Prepare data (fallback to sample data) -----
            DataTable dataTable = new DataTable();

            // Sample data creation – replace with real DB logic if needed
            dataTable.Columns.Add("ID", typeof(int));
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Rows.Add(1, "Alice");
            dataTable.Rows.Add(2, "Bob");

            // Import the DataTable into the worksheet starting at cell A1 (row 0, column 0)
            // Manually write headers
            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                sheet.Cells[0, col].PutValue(dataTable.Columns[col].ColumnName);
            }

            // Write data rows
            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    sheet.Cells[row + 1, col].PutValue(dataTable.Rows[row][col]);
                }
            }

            // ----- Manually trigger calculation -----
            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook to a file with error handling
            try
            {
                string outputPath = "Result.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
