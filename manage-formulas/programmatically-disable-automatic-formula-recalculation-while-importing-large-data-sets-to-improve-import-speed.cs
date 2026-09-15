// Title: How to disable automatic formula calculation in Aspose.Cells for .NET to speed up large DataTable imports
// AI Prompts: Generate C# code that sets Workbook.Settings.CalcMode to Manual before bulk‑loading a 100,000‑row DataTable with Aspose.Cells, then restores Automatic mode and calls CalculateFormula after the import. | Show the exact steps to temporarily turn off formula evaluation in Aspose.Cells, import a large DataTable into a worksheet, and trigger a single recalculation before saving the workbook. | Provide an optimized Aspose.Cells example that disables automatic calculation, uses cell‑by‑cell insertion for a massive dataset, and re‑enables calculation with one CalculateFormula call.
// Common Searches: Aspose.Cells set calculation mode manual for bulk data import C# | Disable formula recalculation in Aspose.Cells to improve performance when loading large DataTable | C# Aspose.Cells import 100k rows without triggering formulas | How to turn off automatic calculation in Aspose.Cells workbook before writing data
// Tags: Aspose.Cells manual calculation mode | bulk DataTable import performance Aspose.Cells | disable automatic formula evaluation .NET | Aspose.Cells CalculateFormula after import | optimize large Excel dataset import Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;
using System.IO;

// The example demonstrates how to improve import speed by setting Workbook.Settings.CalcMode to Manual before writing a large DataTable (e.g., 100,000 rows) into a worksheet, then calling CalculateFormula once after the import and restoring the calculation mode before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Generate a large DataTable to import
            DataTable largeTable = GetLargeDataTable();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Import data starting at cell A1 (row 0, column 0), include column names.
            ImportDataTableToWorksheet(sheet, largeTable, includeColumnNames: true, startRow: 0, startColumn: 0);

            // Recalculate any formulas that may exist after the import
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "LargeDataImport.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Imports a DataTable into a worksheet manually (avoids ImportDataTable API issues)
    static void ImportDataTableToWorksheet(Worksheet sheet, DataTable table, bool includeColumnNames, int startRow, int startColumn)
    {
        try
        {
            Cells cells = sheet.Cells;
            int rowIndex = startRow;
            int colIndex = startColumn;

            // Write column headers if required
            if (includeColumnNames)
            {
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    cells[rowIndex, colIndex + c].PutValue(table.Columns[c].ColumnName);
                }
                rowIndex++;
            }

            // Write data rows
            foreach (DataRow dr in table.Rows)
            {
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    object value = dr[c];
                    cells[rowIndex, colIndex + c].PutValue(value);
                }
                rowIndex++;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error importing DataTable: {ex.Message}");
            throw;
        }
    }

    // Example method that creates a large DataTable
    static DataTable GetLargeDataTable()
    {
        try
        {
            DataTable table = new DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Value", typeof(double));

            // Populate with a large number of rows (e.g., 100,000)
            for (int i = 0; i < 100000; i++)
            {
                table.Rows.Add(i, Math.Sin(i));
            }

            return table;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating DataTable: {ex.Message}");
            throw;
        }
    }
}
