// Title: How to delete a specific column from an Excel ListObject table using Aspose.Cells for .NET (C#)
// Common Searches: C# code to remove a column by name from an Aspose.Cells ListObject | Aspose.Cells delete column from Excel table without affecting other data | remove unwanted column from worksheet table using Aspose.Cells .NET | how to drop a column from an Excel table programmatically with Aspose.Cells | Aspose.Cells ListColumns.RemoveAt example for column removal
// Tags: Aspose.Cells ListObject column deletion | C# Aspose.Cells remove table column | Excel table structure modification Aspose.Cells | ListColumns.RemoveAt method Aspose.Cells | Aspose.Cells .NET column removal example

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The sample loads an existing workbook, accesses the first worksheet's first ListObject table, finds the column named "UnwantedColumn", removes it using ListColumns.RemoveAt, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables found on the worksheet.");
                return;
            }
            ListObject table = sheet.ListObjects[0];

            // Name of the column to remove
            string unwantedColumnName = "UnwantedColumn";

            // Locate the column index within the table
            int columnIndex = -1;
            for (int i = 0; i < table.ListColumns.Count; i++)
            {
                if (table.ListColumns[i].Name.Equals(unwantedColumnName, StringComparison.OrdinalIgnoreCase))
                {
                    columnIndex = i;
                    break;
                }
            }

            // Remove the column if found
            if (columnIndex >= 0)
            {
                table.ListColumns.RemoveAt(columnIndex);
                Console.WriteLine($"Column '{unwantedColumnName}' removed successfully.");
            }
            else
            {
                Console.WriteLine($"Column '{unwantedColumnName}' not found in the table.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
