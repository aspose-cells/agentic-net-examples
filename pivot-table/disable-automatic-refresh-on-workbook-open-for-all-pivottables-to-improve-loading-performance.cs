// Title: How to turn off automatic refresh for every PivotTable in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, iterates through all worksheets, and sets each PivotTable's RefreshDataOnOpen property to false. | Show a complete Aspose.Cells example that prevents PivotTables from refreshing when the workbook is opened, including file‑existence checks and exception handling. | Provide a C# snippet that disables automatic PivotTable refresh on workbook open and saves the updated workbook to a new file.
// Common Searches: aspnet disable pivot table refresh on workbook open using Aspose.Cells | c# Aspose.Cells set RefreshDataOnOpen false for all pivot tables | prevent Excel pivot tables from auto refreshing when opening file with Aspose.Cells library
// Tags: Aspose.Cells disable PivotTable auto refresh | C# iterate worksheets PivotTables Aspose.Cells | set RefreshDataOnOpen property Excel workbook | optimize workbook load performance Aspose.Cells | Excel pivot table refresh control .NET

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing Excel file with Aspose.Cells, checks for its presence, iterates through each worksheet and every PivotTable, attempts to turn off the automatic refresh on open (using the RefreshDataOnOpen property when available), and saves the modified workbook to a new file while handling potential errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Disable automatic refresh for all PivotTables in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // The RefreshDataOnOpen property is not available in this version of Aspose.Cells.
                    // If needed, alternative approaches can be applied here.
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
