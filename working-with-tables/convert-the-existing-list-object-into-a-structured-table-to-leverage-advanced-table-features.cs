// Title: Add a ListObject (structured table) to an existing worksheet from a cell range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create structured table from a range of cells | how to add a ListObject with header row in an existing Excel file using Aspose.Cells | apply built‑in table style to a ListObject in C# Aspose.Cells example | convert Excel range to table and save workbook with Aspose.Cells .NET | C# code to add a ListObject and set TableStyleMedium9 in Aspose.Cells
// Tags: Aspose.Cells ListObject creation | C# structured table insertion | Aspose.Cells built‑in table styling | header row activation ListObject | saving workbook after table addition

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The sample loads an existing workbook, defines a cell range, adds a ListObject (structured table) with a header row over that range, applies a built‑in table style, and saves the modified workbook to a new file.
class Program
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
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range that currently holds the list data (zero‑based indices)
            int firstRow = 0;          // Row 1
            int firstColumn = 0;       // Column A
            int totalRows = 10;        // Number of rows in the list
            int totalColumns = 4;      // Number of columns in the list

            // Add a ListObject (structured table) over the defined range
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true); // true = has header row

            ListObject table = sheet.ListObjects[tableIndex];

            // Apply a built‑in table style (optional)
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Enable header row
            table.ShowHeaderRow = true;

            // NOTE: ShowTotalRow and TotalRowRange are not available in older Aspose.Cells versions.
            // If needed, you can manually add a total row after the data range.

            // Save the workbook with the new structured table
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
