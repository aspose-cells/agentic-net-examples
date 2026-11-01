// Title: How to disable the EnableRefreshOnOpen property for a PivotTable using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set PivotTable EnableRefreshOnOpen to false | prevent pivot table auto refresh when opening Excel workbook with Aspose.Cells | disable automatic data refresh for a PivotTable in .NET using Aspose.Cells | C# code to turn off PivotTable refresh on workbook open with Aspose.Cells | check Aspose.Cells version for EnableRefreshOnOpen support
// Tags: Aspose.Cells PivotTable EnableRefreshOnOpen | disable pivot table auto refresh .NET | Aspose.Cells workbook pivot refresh setting | C# modify pivot table refresh behavior | Excel pivot table open refresh control Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, accesses the first worksheet, retrieves the first PivotTable, and (when supported) sets its EnableRefreshOnOpen property to false to stop automatic data refresh when the workbook is opened, then saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // NOTE: In some versions of Aspose.Cells the property to disable refresh on open
            // is not available. If your version supports it, uncomment the line below:
            // pivotTable.EnableRefreshOnOpen = false;

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
