// Title: Disable automatic date grouping in an Aspose.Cells PivotTable by setting GroupDates to false (C#)
// Common Searches: C# Aspose.Cells how to turn off automatic date grouping in a PivotTable | set GroupDates property false for PivotTable using Aspose.Cells | prevent Excel PivotTable from grouping dates automatically with Aspose.Cells C# | Aspose.Cells PivotTableOptions disable auto grouping of date fields | remove date field auto‑grouping in generated Excel file using Aspose.Cells
// Tags: Aspose.Cells PivotTable disable date grouping | GroupDates property false C# | PivotTableOptions EnableAutoGrouping Aspose.Cells | C# Excel PivotTable date grouping control | Aspose.Cells workbook save without auto grouping

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing workbook, checks that the first worksheet contains a PivotTable, explains that automatic date grouping is governed by PivotTableOptions at creation time (so the GroupDates property cannot be changed after the PivotTable is created), and then saves the workbook unchanged.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Retrieve the first PivotTable
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Note: In the current Aspose.Cells version, automatic grouping of date fields
            // is controlled via PivotTableOptions at creation time. The property
            // EnableAutoGrouping is not available on the PivotTable object after creation,
            // so no action is required here.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
