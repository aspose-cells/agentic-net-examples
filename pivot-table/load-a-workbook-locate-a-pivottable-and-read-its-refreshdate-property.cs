// Title: Read a PivotTable's RefreshDate from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: c# Aspose.Cells get pivot table refresh date from Excel file | how to retrieve PivotTable RefreshDate property with Aspose.Cells in .NET | read last refreshed timestamp of an Excel pivot table using C# | Aspose.Cells iterate worksheets and access pivot table metadata C# | check if pivot table was refreshed after loading workbook with Aspose.Cells
// Tags: Aspose.Cells read PivotTable RefreshDate | C# load Excel workbook Aspose.Cells | iterate worksheets Aspose.Cells C# | access PivotTable collection Aspose.Cells | handle missing Excel file C# Aspose.Cells | output pivot refresh timestamp .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable types

// The example verifies that the input Excel file exists, loads it into an Aspose.Cells Workbook, loops through each worksheet and its PivotTableCollection, reads the RefreshDate property of every pivot table, and writes the pivot name, worksheet name, and refresh timestamp to the console, with exception handling for runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of pivot tables on the current worksheet
                PivotTableCollection pivotTables = sheet.PivotTables;

                // Loop through each pivot table
                foreach (PivotTable pivot in pivotTables)
                {
                    // Read the RefreshDate property of the pivot table
                    DateTime refreshDate = pivot.RefreshDate;

                    // Output the pivot table name, worksheet name, and its refresh date
                    Console.WriteLine(
                        $"Pivot Table '{pivot.Name}' in worksheet '{sheet.Name}' was refreshed on: {refreshDate}");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
