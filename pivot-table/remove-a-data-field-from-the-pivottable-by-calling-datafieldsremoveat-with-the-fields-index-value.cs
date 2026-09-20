// Title: How to remove a data field from a PivotTable in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, checks for a PivotTable, validates a zero‑based data field index, removes that field using DataFields.RemoveAt, and saves the workbook. | Create a robust example that demonstrates deleting the first data field from a PivotTable, including error handling for missing files and invalid indexes. | Generate a snippet that iterates over PivotTable.DataFields and removes a specific field by its position with Aspose.Cells.
// Common Searches: aspnet remove specific data field from pivot table using Aspose.Cells | c# Aspose.Cells delete pivot table data field by index example | how to programmatically remove a data field from an Excel pivot table with Aspose.Cells | validate pivot table data field index before removal Aspose.Cells C# | remove first data field from pivot table and save workbook Aspose.Cells
// Tags: Aspose.Cells PivotTable DataFields.RemoveAt | C# remove pivot table data field | Excel pivot table field deletion Aspose | validate data field index Aspose.Cells | save workbook after pivot modification C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, accesses the first worksheet's first PivotTable, checks that a specified zero‑based data field index is valid, removes that data field with DataFields.RemoveAt, and saves the modified workbook to a new file, handling missing files and invalid indexes gracefully.
class RemovePivotDataField
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists before loading
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Get the first PivotTable
            PivotTable pivotTable = sheet.PivotTables[0];

            // Index of the data field to remove (zero‑based)
            int dataFieldIndex = 0;

            // Validate the index against existing data fields
            if (dataFieldIndex < 0 || dataFieldIndex >= pivotTable.DataFields.Count)
            {
                Console.WriteLine("Invalid data field index.");
                return;
            }

            // Remove the specified data field
            pivotTable.DataFields.RemoveAt(dataFieldIndex);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
