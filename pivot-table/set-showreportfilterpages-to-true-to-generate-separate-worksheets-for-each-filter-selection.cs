// Title: Create separate worksheets for each pivot table report filter value using ShowReportFilterPages in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# generate a worksheet for each pivot table report filter item | How to invoke ShowReportFilterPages dynamically with Aspose.Cells .NET | Create separate Excel sheets per pivot filter selection using Aspose.Cells API | C# example handling missing ShowReportFilterPages method in older Aspose.Cells versions | Pivot table report filter pages generation with Aspose.Cells for .NET
// Tags: ShowReportFilterPages Aspose.Cells C# | pivot table report filter pages generation | dynamic method invocation Aspose.Cells | separate worksheet per filter value .NET | Aspose.Cells pivot table worksheet duplication

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an Excel workbook, checks for a pivot table, and uses dynamic invocation to call ShowReportFilterPages for a specified report filter field, creating a new worksheet for each filter value. It gracefully handles cases where the method is unavailable and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains a pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Name of the report filter field for which separate pages should be created
            string reportFilterField = "YourFilterFieldName";

            // Attempt to generate separate worksheets for each selection of the specified report filter.
            // The ShowReportFilterPages method may not be available in all versions; use dynamic invocation to avoid compile‑time errors.
            try
            {
                ((dynamic)pivotTable).ShowReportFilterPages(reportFilterField);
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                Console.WriteLine("ShowReportFilterPages method is not available in the current Aspose.Cells version.");
            }

            // Save the workbook with the newly created worksheets (if any)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
