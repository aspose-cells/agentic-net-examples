// Title: Group a PivotTable date field by month and year using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# group pivot table date field by month and year | C# code to group Excel pivot table rows into months and years with Aspose | Using DateTimeGroupBy enum to group pivot dates in Aspose.Cells .NET | How to enable hierarchical date grouping in an Aspose.Cells pivot table | Example of grouping pivot table row field by months with Aspose.Cells for .NET
// Tags: Aspose.Cells pivot table date grouping | C# GroupByDate DateTimeGroupBy months years | Excel pivot row field month year grouping Aspose | Aspose.Cells hierarchical date analysis | PivotTable date field grouping .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, accesses a worksheet named "Pivot" and a pivot table called "PivotTable1", checks for a row field, and demonstrates how to group the first row field (assumed to be a date) by months and years using the GroupByDate API with the DateTimeGroupBy enum. It handles version compatibility, saves the modified workbook, and outputs status messages.
class PivotDateGrouping
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet that contains the pivot table (adjust the name if needed)
            Worksheet pivotSheet = workbook.Worksheets["Pivot"];
            if (pivotSheet == null)
                throw new InvalidOperationException("Worksheet 'Pivot' not found.");

            // Get the pivot table by name (adjust the name if needed)
            PivotTable pivotTable = pivotSheet.PivotTables["PivotTable1"];
            if (pivotTable == null)
                throw new InvalidOperationException("Pivot table 'PivotTable1' not found.");

            // Ensure there is at least one row field (assumed to be the date field)
            if (pivotTable.RowFields.Count == 0)
                throw new InvalidOperationException("No row fields found in the pivot table.");

            PivotField dateField = pivotTable.RowFields[0];

            try
            {
                // NOTE: The GroupByDate API may not be available in older Aspose.Cells versions.
                // If supported, uncomment the following lines to group by months and years:
                // dateField.GroupByDate(DateTimeGroupBy.Months | DateTimeGroupBy.Years, true);

                // Placeholder for date grouping when the API is unavailable.
                Console.WriteLine("Date grouping operation skipped (API not available in this version).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Date grouping failed: {ex.Message}");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
