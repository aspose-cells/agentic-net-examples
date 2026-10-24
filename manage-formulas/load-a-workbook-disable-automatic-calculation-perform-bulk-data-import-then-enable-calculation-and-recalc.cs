// Title: Bulk import a DataTable into an existing Excel workbook and force full formula recalculation using Aspose.Cells for .NET
// Common Searches: how to write a DataTable to an existing Excel file with Aspose.Cells in C# | Aspose.Cells bulk data import into worksheet without triggering calculation | recalculate all formulas after inserting rows in an Excel workbook using Aspose.Cells | C# load workbook, add DataTable rows, then call CalculateFormula | disable automatic calculation during data import with Aspose.Cells C#
// Tags: bulk import DataTable to Excel worksheet Aspose.Cells | calculate all formulas workbook.CalculateFormula Aspose.Cells | load existing workbook and write cells C# Aspose.Cells | disable automatic calculation mode Aspose.Cells | save modified workbook Output.xlsx Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads an existing Input.xlsx workbook, creates a DataTable with product information, writes the column headers and rows to the first worksheet starting at cell A1, forces a full recalculation of all formulas using workbook.CalculateFormula(), and saves the updated file as Output.xlsx, handling errors gracefully.
class Program
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare bulk data (example DataTable)
            DataTable table = new DataTable();
            table.Columns.Add("Product", typeof(string));
            table.Columns.Add("Quantity", typeof(int));
            table.Columns.Add("Price", typeof(double));

            // Populate the DataTable with sample rows
            table.Rows.Add("Apple", 120, 0.5);
            table.Rows.Add("Banana", 80, 0.3);
            table.Rows.Add("Orange", 150, 0.4);

            // Perform bulk import into the first worksheet starting at cell A1 (row 0, column 0)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Write header
            for (int col = 0; col < table.Columns.Count; col++)
            {
                cells[0, col].PutValue(table.Columns[col].ColumnName);
            }

            // Write data rows
            for (int row = 0; row < table.Rows.Count; row++)
            {
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    cells[row + 1, col].PutValue(table.Rows[row][col]);
                }
            }

            // Force a full recalculation of all formulas
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
