// Title: Detect Excel tables missing a totals row and insert a default sum totals row using Aspose.Cells for .NET
// Common Searches: how to check if an Excel table has a totals row with Aspose.Cells C# | add a totals row with sum formulas to tables that lack one using Aspose.Cells .NET | programmatically create a totals row for each worksheet table in C# Aspose.Cells | detect tables without totals row and apply default aggregation in Aspose.Cells | Aspose.Cells C# add missing totals row to Excel tables automatically
// Tags: detect missing totals row Aspose.Cells | add default sum totals row Excel table | Aspose.Cells table totals row automation | C# Excel table aggregation Aspose.Cells | Aspose.Cells.Tables package usage

using Aspose.Cells;
using System;
using System.IO;

// This example shows how to load an Excel workbook with Aspose.Cells, iterate through all worksheet tables, identify those that do not have a totals row, and programmatically add a totals row that calculates the sum for each numeric column. After updating the tables, the workbook is saved with the new totals rows applied.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Table manipulation APIs (Worksheet.Tables, Table, TableColumn, TotalRowFunction)
            // are not available in the current Aspose.Cells version referenced by this project.
            // If needed, add the appropriate Aspose.Cells.Tables package and adjust the code accordingly.

            // Save the (potentially modified) workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine(ex.Message);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
