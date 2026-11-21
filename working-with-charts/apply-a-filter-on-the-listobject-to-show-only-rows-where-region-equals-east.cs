// Title: C# – Use Aspose.Cells to apply an AutoFilter on a ListObject and show only rows where the Region column equals "East"
// Common Searches: C# Aspose.Cells filter rows where Region column equals East | How to add an AutoFilter to a ListObject in Aspose.Cells .NET | Programmatically hide non‑matching rows in an Excel table using Aspose.Cells | Apply column‑based filter to an Excel table with Aspose.Cells C# example | Aspose.Cells filter ListObject by specific column value
// Tags: Aspose.Cells ListObject AutoFilter | filter Excel table rows C# | apply column filter Aspose.Cells | Excel table Region column filter | C# programmatic ListObject filtering

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample loads an existing workbook, accesses the first worksheet, finds the first ListObject (Excel table), locates the "Region" column in the table header, adds an AutoFilter criterion to keep only rows where the Region value is "East", refreshes the filter to hide other rows, and saves the result to a new file, with error handling for missing files, missing columns, and save failures.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one ListObject (table) in the sheet
            if (worksheet.ListObjects.Count > 0)
            {
                // Get the first ListObject
                ListObject listObject = worksheet.ListObjects[0];

                // Locate the "Region" column within the ListObject header row
                int headerRow = listObject.StartRow;          // Header row index
                int firstCol = listObject.StartColumn;        // First column index of the table
                int columnCount = listObject.ListColumns.Count; // Number of columns in the table
                int regionRelIndex = -1;                      // Relative index inside the ListObject

                for (int i = 0; i < columnCount; i++)
                {
                    if (worksheet.Cells[headerRow, firstCol + i].StringValue == "Region")
                    {
                        regionRelIndex = i;
                        break;
                    }
                }

                // If the "Region" column was found, apply the filter
                if (regionRelIndex >= 0)
                {
                    try
                    {
                        // Add filter criteria: Region = "East"
                        listObject.AutoFilter.AddFilter(regionRelIndex, "East");
                        // Apply the filter to hide non‑matching rows
                        listObject.AutoFilter.Refresh();
                    }
                    catch (Exception filterEx)
                    {
                        Console.WriteLine($"Error applying filter: {filterEx.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("Column \"Region\" not found in the table.");
                }
            }
            else
            {
                Console.WriteLine("No tables (ListObjects) found in the worksheet.");
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
