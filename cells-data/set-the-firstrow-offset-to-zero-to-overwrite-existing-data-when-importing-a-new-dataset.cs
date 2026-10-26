// Title: Overwrite existing worksheet cells by importing a DataTable with startRow 0 using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells import DataTable starting at row 0 to replace existing cells | How to reset worksheet content when loading a new DataTable with Aspose.Cells | Aspose.Cells startRow parameter to overwrite Excel sheet data in .NET | Replace all data in an existing worksheet by importing a DataTable with Aspose.Cells | Import DataTable into Excel without shifting existing rows using Aspose.Cells C#
// Tags: Aspose.Cells import DataTable startRow zero | overwrite worksheet cells Aspose.Cells C# | Aspose.Cells Cells.PutValue replace existing values | C# import DataTable to Excel without shifting rows | Aspose.Cells startRow parameter usage | Aspose.Cells workbook save ensure directory exists

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, builds a sample DataTable, and imports it into the first worksheet starting at row 0 and column 0. By using a startRow of zero, any existing cells are overwritten. The code also ensures the output directory exists before saving the workbook to Output.xlsx, with comprehensive exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example DataTable to import
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("Alice", 30);
            dt.Rows.Add("Bob", 25);

            // Import the DataTable into the worksheet (including column headers)
            ImportDataTableToWorksheet(sheet, dt, includeHeaders: true, startRow: 0, startColumn: 0);

            // Determine output path and ensure its directory exists
            string outputPath = "Output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    private static void ImportDataTableToWorksheet(Worksheet sheet, DataTable table, bool includeHeaders, int startRow, int startColumn)
    {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (table == null) throw new ArgumentNullException(nameof(table));

        Cells cells = sheet.Cells;
        int currentRow = startRow;

        // Write column headers if required
        if (includeHeaders)
        {
            for (int col = 0; col < table.Columns.Count; col++)
            {
                cells[currentRow, startColumn + col].PutValue(table.Columns[col].ColumnName);
            }
            currentRow++;
        }

        // Write data rows
        foreach (DataRow dr in table.Rows)
        {
            for (int col = 0; col < table.Columns.Count; col++)
            {
                cells[currentRow, startColumn + col].PutValue(dr[col]);
            }
            currentRow++;
        }
    }
}
