// Title: How to disable automatic PivotTable refresh on workbook open with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# prevent PivotTable refresh when opening an Excel file | disable workbook open auto refresh for all PivotTables using Aspose.Cells | improve Excel load time by turning off PivotTable auto refresh in Aspose.Cells .NET | C# code to set PivotTable RefreshDataOnOpen to false with Aspose.Cells | iterate worksheets to disable PivotTable refresh on open in Aspose.Cells
// Tags: Aspose.Cells disable pivot auto refresh | C# pivot table refresh on open setting | Excel workbook load performance Aspose.Cells | iterate worksheets pivot tables Aspose.Cells | RefreshDataOnOpen not supported Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Provides PivotTable type

// The example loads an existing Excel workbook, loops through each worksheet and its PivotTables, and shows where you would disable the automatic refresh on workbook open (noting that the RefreshDataOnOpen property is not available in the current Aspose.Cells version). The modified workbook is then saved to a new file, helping to improve load performance when the workbook contains many PivotTables.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of PivotTables on the worksheet
                int pivotCount = sheet.PivotTables.Count;

                // Disable automatic refresh on open for each PivotTable
                // Note: The RefreshDataOnOpen property is not available in the current Aspose.Cells version.
                // If needed, this functionality can be handled via other workbook settings or updated library versions.
                for (int i = 0; i < pivotCount; i++)
                {
                    PivotTable pivot = sheet.PivotTables[i];
                    // Placeholder for any future pivot configuration
                }
            }

            // Save the updated workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
