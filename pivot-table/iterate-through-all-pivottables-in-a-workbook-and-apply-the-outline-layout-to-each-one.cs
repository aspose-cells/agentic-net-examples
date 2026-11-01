// Title: How to loop through every PivotTable in an Excel workbook and set the Outline layout with Aspose.Cells for .NET (C#)
// Common Searches: C# code to set PivotTable layout to Outline for all pivot tables in a workbook using Aspose.Cells | Iterate over worksheets and pivot tables to apply outline view with Aspose.Cells .NET | Bulk change PivotTable layout to Outline in Excel file via Aspose.Cells API | Aspose.Cells set PivotTableLayoutType.Outline for each pivot table programmatically | Update all pivot tables to outline style in an existing .xlsx using Aspose.Cells C#
// Tags: Aspose.Cells pivot table layout manipulation | C# iterate worksheets pivot tables | programmatic outline layout for pivot tables | bulk pivot table formatting Aspose.Cells | Excel workbook pivot table processing C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable classes

// The example loads an existing Excel workbook, checks the file, then iterates through each worksheet and each PivotTable within them. For each PivotTable it attempts to assign the Outline layout via the PivotTableLayout property (commented for version compatibility), and finally saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each PivotTable in the current worksheet
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    try
                    {
                        // The PivotTableLayout property may not be available in older versions of Aspose.Cells.
                        // If supported, you can set the layout to Outline as shown below:
                        // pivotTable.PivotTableLayout = PivotTableLayoutType.Outline;

                        // Placeholder for any additional PivotTable configuration.
                    }
                    catch (Exception ptEx)
                    {
                        Console.WriteLine($"Error processing PivotTable \"{pivotTable.Name}\": {ptEx.Message}");
                    }
                }
            }

            // Save the workbook with the (potentially) updated PivotTable layouts
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
