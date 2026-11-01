// Title: How to delete a specific column from an Excel ListObject table using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove column from Excel table by column name | Delete unwanted column from ListObject without affecting other data in .NET | C# code to drop a column from an Excel table using Aspose.Cells | How to remove a column from a worksheet table and save workbook with Aspose.Cells | Remove column from Excel table programmatically in C# Aspose.Cells example
// Tags: Aspose.Cells ListObject column deletion C# | Excel table column removal using Aspose | C# ListColumns.RemoveAt method Aspose.Cells | preserve other columns when dropping table column | save workbook after modifying table structure Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an XLSX workbook, accesses the first worksheet's first ListObject (Excel table), locates a column named "UnwantedColumn", removes that column via the ListColumns collection while leaving remaining data intact, and saves the updated workbook to a new file, with error handling for missing files and other exceptions.
class RemoveColumnFromTable
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (worksheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables found on the worksheet.");
                return;
            }

            // Assume the first table is the target
            ListObject table = worksheet.ListObjects[0];

            // Name of the column to remove
            string unwantedColumnName = "UnwantedColumn";

            // Find the zero‑based index of the column within the table
            int columnIndex = -1;
            for (int i = 0; i < table.ListColumns.Count; i++)
            {
                if (table.ListColumns[i].Name.Equals(unwantedColumnName, StringComparison.OrdinalIgnoreCase))
                {
                    columnIndex = i;
                    break;
                }
            }

            // Delete the column if it exists
            if (columnIndex != -1)
            {
                // Remove the column from the table's ListColumns collection
                table.ListColumns.RemoveAt(columnIndex);
                Console.WriteLine($"Column \"{unwantedColumnName}\" removed successfully.");
            }
            else
            {
                Console.WriteLine($"Column \"{unwantedColumnName}\" not found in the table.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
