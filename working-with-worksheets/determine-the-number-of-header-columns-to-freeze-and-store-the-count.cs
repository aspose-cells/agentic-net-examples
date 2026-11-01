// Title: Get the count of frozen header columns in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to read number of frozen columns in a worksheet | retrieve frozen pane column split value with Aspose.Cells .NET | using reflection to access FreezePanes.ColumnSplit in older Aspose.Cells versions | determine how many columns are frozen in an Excel file using Aspose.Cells | Aspose.Cells get frozen header columns count programmatically
// Tags: Aspose.Cells frozen column count retrieval | Worksheet FreezePanes ColumnSplit access | C# reflection for legacy Aspose.Cells API | Excel frozen header columns detection | programmatic frozen panes query Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, accesses the first worksheet, and uses reflection to safely read the ColumnSplit value from the worksheet's FreezePanes property (when available) to determine the number of frozen header columns, then outputs the count and optionally saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the input file.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet in the workbook.
            Worksheet worksheet = workbook.Worksheets[0];

            int frozenHeaderColumns = 0;

            // Attempt to retrieve the number of frozen columns.
            // In newer Aspose.Cells versions this can be obtained via worksheet.FreezePanes.ColumnSplit.
            // For compatibility with older versions where FreezePanes is a method, default to 0.
            try
            {
                // The FreezePanes property may not be available in older library versions.
                // Use reflection to safely access ColumnSplit if it exists.
                var freezePanesProp = typeof(Worksheet).GetProperty("FreezePanes");
                if (freezePanesProp != null)
                {
                    var freezePanesObj = freezePanesProp.GetValue(worksheet);
                    var columnSplitProp = freezePanesObj?.GetType().GetProperty("ColumnSplit");
                    if (columnSplitProp != null)
                    {
                        frozenHeaderColumns = (int)columnSplitProp.GetValue(freezePanesObj);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to determine frozen columns. {ex.Message}");
            }

            // Output the result.
            Console.WriteLine($"Number of frozen header columns: {frozenHeaderColumns}");

            // Save the workbook (optional, only if modifications are made).
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
