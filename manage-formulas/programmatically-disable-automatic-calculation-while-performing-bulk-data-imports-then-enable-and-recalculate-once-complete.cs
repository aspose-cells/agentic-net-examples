// Title: How to disable automatic formula calculation during bulk DataTable import and recalculate after import with Aspose.Cells for .NET
// AI Prompts: Write C# that disables automatic formula evaluation, loads a DataTable into the first worksheet, then re‑enables evaluation and calls CalculateFormula. | Show example code using Aspose.Cells to turn off formula calculation while bulk‑loading rows and then trigger a full recalculation after the load.
// Common Searches: Aspose.Cells set calculation mode to manual for bulk data import C# | Improve performance when inserting thousands of rows with formulas using Aspose.Cells | Re‑enable automatic calculation after bulk write in Aspose.Cells .NET | Disable formula recalculation during DataTable export to Excel with Aspose.Cells | Calculate all formulas after disabling automatic calculation in Aspose.Cells workbook
// Tags: manual calculation mode Aspose.Cells | DataTable bulk load Aspose.Cells | full formula recalculation Aspose.Cells | performance optimization Excel export .NET | disable automatic formula evaluation Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The program creates a new workbook, writes column headers and rows from a DataTable into the first worksheet, forces a full formula recalculation with CalculateFormula, and saves the workbook as BulkImportResult.xlsx.
class BulkImportExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (or add a new one if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // Bulk data import section
            // Replace this placeholder with your actual data source (e.g., DataTable, array, etc.)
            // -----------------------------------------------------------------
            DataTable data = GetSampleData(); // Example method returning a DataTable

            // Write column headers
            for (int col = 0; col < data.Columns.Count; col++)
            {
                sheet.Cells[0, col].PutValue(data.Columns[col].ColumnName);
            }

            // Write data rows
            for (int row = 0; row < data.Rows.Count; row++)
            {
                for (int col = 0; col < data.Columns.Count; col++)
                {
                    sheet.Cells[row + 1, col].PutValue(data.Rows[row][col]);
                }
            }
            // -----------------------------------------------------------------

            // Force a full recalculation of all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook to a file
            string outputPath = "BulkImportResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Sample method to generate a DataTable for demonstration purposes
    private static DataTable GetSampleData()
    {
        DataTable table = new DataTable();
        table.Columns.Add("Product", typeof(string));
        table.Columns.Add("Quantity", typeof(int));
        table.Columns.Add("Price", typeof(double));

        table.Rows.Add("Apple", 120, 0.5);
        table.Rows.Add("Banana", 85, 0.3);
        table.Rows.Add("Cherry", 200, 0.2);

        return table;
    }
}
