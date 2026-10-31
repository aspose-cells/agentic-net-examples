// Title: Enable row grand totals in an Excel PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: how to enable row grand totals in an Excel PivotTable using Aspose.Cells C# | Aspose.Cells C# set RowGrandTotals property for existing PivotTable | programmatically display row totals at the bottom of a PivotTable with .NET
// Tags: Aspose.Cells enable row grand totals | C# PivotTable RowGrandTotals API | modify existing Excel workbook PivotTable | reflection based property setting Aspose.Cells | Excel PivotTable totals configuration .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, checks for a PivotTable on the first worksheet, and uses reflection to set the PivotTable's RowGrandTotals (ShowGrandTotalsForRows) property to true when available, then saves the updated file.
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
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or specify the appropriate index/name)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure that at least one PivotTable exists on the worksheet
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found on the first worksheet.");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved without changes to: {outputPath}");
                return;
            }

            // Retrieve the first PivotTable
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Enable row grand totals if the property is available.
            // In some older Aspose.Cells versions the RowGrandTotals property may not exist.
            // If available, set it to true; otherwise, this step is skipped.
            try
            {
                // Attempt to set RowGrandTotals using reflection to maintain compatibility across versions
                var prop = typeof(PivotTable).GetProperty("RowGrandTotals");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(pivotTable, true);
                }
                else
                {
                    Console.WriteLine("RowGrandTotals property not available in this Aspose.Cells version.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to set RowGrandTotals: {ex.Message}");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions (e.g., Aspose.Cells specific errors)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
