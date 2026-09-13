// Title: Dynamically assign TableToRangeOptions.LastRow from a ListObject’s row count before converting an Excel table to a range with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook, checks for the first ListObject, calculates its data row count, sets TableToRangeOptions.LastRow to that count minus one, and converts the table to a range. | Show how to safely verify file existence and the presence of tables, then use Aspose.Cells TableToRangeOptions to convert a table to a range with a dynamically computed last row. | Provide a snippet that demonstrates retrieving ListObject.Rows.Count, configuring TableToRangeOptions.LastRow, and saving the workbook after the conversion.
// Common Searches: aspnet aspose.cells set TableToRangeOptions.LastRow based on ListObject row count | convert excel table to range with dynamic last row using Aspose.Cells C# | how to get number of data rows in a ListObject with Aspose.Cells | Aspose.Cells TableToRangeOptions LastRow example for .NET | C# check for tables before converting to range in Aspose.Cells
// Tags: runtime last row calculation Aspose.Cells | listobject row count to range conversion | c# aspose.cells convert table with options | verify table existence before conversion | excel workbook table to range .net

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The program loads "input.xlsx", confirms the file and that the first worksheet contains a table, obtains the first ListObject, computes its data row count, sets TableToRangeOptions.LastRow to rowCount‑1, converts the table to a range, and saves the result as "output.xlsx" with basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table
            if (worksheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables (ListObjects) found in the worksheet.");
                return;
            }

            // Access the first ListObject (table) on the worksheet
            ListObject table = worksheet.ListObjects[0];

            // Determine the number of rows in the table's data range
            int rowCount = table.DataRange.RowCount;

            // Create conversion options and set the last row dynamically (zero‑based index)
            TableToRangeOptions options = new TableToRangeOptions
            {
                LastRow = rowCount - 1
            };

            // Convert the table to a range using the defined options
            table.ConvertToRange(options);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
