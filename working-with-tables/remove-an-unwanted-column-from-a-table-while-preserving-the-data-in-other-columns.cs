// Title: Remove an unwanted column from an Excel ListObject table using Aspose.Cells for .NET while keeping other data intact
// AI Prompts: Generate C# code that locates a ListObject column by its header name and deletes it with Aspose.Cells, leaving the rest of the table unchanged. | Show how to programmatically drop a specific column from an Excel table (ListObject) in a workbook using Aspose.Cells for .NET. | Create a reusable method that accepts a worksheet and column header, then removes that column from the first table using Aspose.Cells without affecting other columns.
// Common Searches: aspnet remove column from Excel table ListObject Aspose.Cells example | c# delete specific column in an Excel worksheet table using Aspose.Cells | how to drop a table column while preserving other columns in Aspose.Cells for .NET | remove unwanted column from first ListObject in workbook with Aspose.Cells C# | Aspose.Cells delete table column by header name programmatically
// Tags: Aspose.Cells delete ListObject column C# | preserve other columns when dropping Excel table column Aspose.Cells | C# Aspose.Cells delete column by header name | modify worksheet table column removal Aspose.Cells | programmatic column deletion in Excel table .NET

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// // Loads an existing workbook, accesses the first worksheet and its first ListObject, finds the column named 'UnwantedColumn', deletes that column while leaving all other columns untouched, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (change index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Verify that the worksheet contains at least one table (list object)
            if (worksheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables found in the worksheet.");
                return;
            }

            // Get the first table in the worksheet
            ListObject table = worksheet.ListObjects[0];

            // Name of the column to be removed
            string columnNameToRemove = "UnwantedColumn";

            // Locate the column index within the table (0‑based)
            int columnIndex = -1;
            for (int i = 0; i < table.ListColumns.Count; i++)
            {
                if (table.ListColumns[i].Name.Equals(columnNameToRemove, StringComparison.OrdinalIgnoreCase))
                {
                    columnIndex = i;
                    break;
                }
            }

            if (columnIndex >= 0)
            {
                // Remove the column while preserving data in other columns
                table.ListColumns.RemoveAt(columnIndex);
                Console.WriteLine($"Column '{columnNameToRemove}' removed successfully.");
            }
            else
            {
                Console.WriteLine($"Column '{columnNameToRemove}' not found in the table.");
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
