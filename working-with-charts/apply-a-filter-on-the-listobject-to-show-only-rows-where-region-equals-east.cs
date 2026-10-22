// Title: How to filter rows in an Excel ListObject where Region equals "East" using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, finds the first ListObject, and applies an AutoFilter to display only rows where the "Region" column is "East". | Provide a step‑by‑step C# example that uses Aspose.Cells to locate a table column by name, set a filter criterion, and save the filtered workbook.
// Common Searches: asp.net filter excel table rows where column value is East using Aspose.Cells | c# Aspose.Cells apply AutoFilter to ListObject column Region | how to show only rows with Region = East in an Excel table with Aspose.Cells | filter ListObject by column value in Aspose.Cells .NET example
// Tags: Aspose.Cells ListObject AutoFilter C# | Excel table column value filter using Aspose.Cells | Region column filter Aspose.Cells | C# hide rows with Aspose.Cells AutoFilter | Aspose.Cells filter criteria East

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// // Loads "input.xlsx", locates the first ListObject, identifies the "Region" column, applies an AutoFilter to show only rows where Region equals "East", and saves the result as "output.xlsx".
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one ListObject (table)
            if (worksheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No ListObjects (tables) found in the worksheet.");
                return;
            }

            // Get the first ListObject
            ListObject listObject = worksheet.ListObjects[0];

            // Find the column index for the column named "Region"
            int regionColumnIndex = -1;
            for (int i = 0; i < listObject.ListColumns.Count; i++)
            {
                ListColumn col = listObject.ListColumns[i];
                if (string.Equals(col.Name, "Region", StringComparison.OrdinalIgnoreCase))
                {
                    regionColumnIndex = i;
                    break;
                }
            }

            if (regionColumnIndex == -1)
            {
                Console.WriteLine("Column 'Region' not found in the table.");
                return;
            }

            // Apply a filter to show only rows where Region equals "East"
            listObject.AutoFilter.Filter(regionColumnIndex, "East");

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
