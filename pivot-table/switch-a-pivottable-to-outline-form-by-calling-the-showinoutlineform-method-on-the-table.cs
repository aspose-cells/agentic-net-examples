// Title: How to convert an Excel PivotTable to Outline view using Aspose.Cells ShowInOutlineForm in C#
// Common Searches: C# code to set Excel pivot table to outline form with Aspose.Cells | Aspose.Cells ShowInOutlineForm method example for .NET | Programmatically change pivot table layout to outline in a workbook using Aspose.Cells | Switch pivot table display mode to outline using Aspose.Cells for .NET | Load workbook, modify first pivot table to outline view, save with Aspose.Cells C#
// Tags: Aspose.Cells ShowInOutlineForm C# | pivot table outline layout Aspose.Cells | modify pivot table display mode .NET | Excel workbook pivot table outline conversion | programmatic pivot table layout change Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, accesses the first worksheet, retrieves the first PivotTable, switches its layout to outline form by calling ShowInOutlineForm(), and saves the modified workbook to a new file using Aspose.Cells for .NET.
class PivotTableOutlineExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure at least one pivot table exists
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];

            // Switch the pivot table to outline form
            pivotTable.ShowInOutlineForm();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
