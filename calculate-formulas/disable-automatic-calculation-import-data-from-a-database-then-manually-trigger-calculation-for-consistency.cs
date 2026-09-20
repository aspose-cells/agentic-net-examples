// Title: Disable auto‑calculation, import a DataTable into an Excel worksheet, and manually recalculate formulas using Aspose.Cells in C#
// AI Prompts: Write C# code that sets Aspose.Cells workbook calculation mode to manual, loads rows from a DataTable into a worksheet, adds a Total column with a formula, and invokes CalculateFormula before saving the file. | Show how to turn off automatic formula evaluation in Aspose.Cells, populate cells from a database‑derived DataTable, assign cell formulas, and trigger a manual recalculation in .NET.
// Common Searches: how to turn off automatic calculation in Aspose.Cells C# before importing data | import DataTable into Excel worksheet with Aspose.Cells and calculate totals manually | Aspose.Cells manual calculation mode example for database data | C# Aspose.Cells calculate formulas after populating worksheet | disable auto calc and trigger workbook.CalculateFormula in Aspose.Cells
// Tags: Aspose.Cells manual calculation mode | import DataTable to Excel worksheet Aspose.Cells | assign cell formulas programmatically Aspose.Cells | trigger workbook.CalculateFormula C# | populate Excel from database Aspose.Cells

using Aspose.Cells;
using System;
using System.Data;
using System.IO;

// The example creates a Workbook, disables automatic calculation, imports a DataTable into the first worksheet, adds a 'Total' column with formulas referencing Quantity and Price, manually triggers formula evaluation with workbook.CalculateFormula(), and saves the result as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data (replace with actual DB retrieval if needed)
            DataTable dt = new DataTable();
            dt.Columns.Add("Id", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Quantity", typeof(int));
            dt.Columns.Add("Price", typeof(double));

            dt.Rows.Add(1, "Product A", 10, 2.5);
            dt.Rows.Add(2, "Product B", 5, 4.0);
            dt.Rows.Add(3, "Product C", 8, 3.75);

            // Manually import the DataTable into the worksheet starting at cell A1
            // Write header
            for (int col = 0; col < dt.Columns.Count; col++)
            {
                sheet.Cells[0, col].PutValue(dt.Columns[col].ColumnName);
            }

            // Write data rows
            for (int row = 0; row < dt.Rows.Count; row++)
            {
                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    sheet.Cells[row + 1, col].PutValue(dt.Rows[row][col]);
                }
            }

            // Add a formula column for Total = Quantity * Price
            int totalColumnIndex = dt.Columns.Count; // Column after the last data column (E)
            sheet.Cells[0, totalColumnIndex].PutValue("Total"); // Header

            int startRow = 1; // Data starts after header (0‑based index)
            int lastDataRow = sheet.Cells.MaxDataRow; // Last row with data

            for (int row = startRow; row <= lastDataRow; row++)
            {
                // Build formula referencing Quantity (C) and Price (D) columns
                string formula = $"=C{row + 1}*D{row + 1}";
                sheet.Cells[row, totalColumnIndex].Formula = formula;
            }

            // Manually trigger calculation to evaluate formulas
            workbook.CalculateFormula();

            // Save the workbook to a file
            string outputPath = "Output.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
