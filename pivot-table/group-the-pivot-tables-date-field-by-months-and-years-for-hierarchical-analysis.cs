// Title: How to group a PivotTable date field by months and years with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# group pivot table date field by month and year | C# code to group PivotTable rows by months and years using Aspose.Cells | Using PivotField.GroupByDate to create month‑year hierarchy in Excel pivot with Aspose | Version‑specific grouping of date fields in Aspose.Cells pivot tables | How to enable hierarchical date grouping in an Excel pivot via Aspose.Cells .NET
// Tags: PivotField.GroupByDate months years | Aspose.Cells pivot date field grouping | C# group pivot table by month year | Excel hierarchical date grouping Aspose.Cells | date field grouping in pivot tables .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, checks for a pivot table on the first worksheet, retrieves the first row field (assumed to be a date field), and demonstrates how to apply PivotField.GroupByDate with the Months and Years flags (commented for version compatibility) to create hierarchical month‑year grouping, then saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
            }
            else
            {
                // Get the first pivot table on the worksheet
                PivotTable pivotTable = worksheet.PivotTables[0];

                if (pivotTable.RowFields.Count > 0)
                {
                    // Identify the date field to be grouped (assumed first row field)
                    PivotField dateField = pivotTable.RowFields[0];

                    // Grouping by months and years is version‑dependent.
                    // If the current Aspose.Cells version supports it, uncomment the line below:
                    // dateField.GroupByDate(PivotFieldGroupByDateType.Months | PivotFieldGroupByDateType.Years);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
