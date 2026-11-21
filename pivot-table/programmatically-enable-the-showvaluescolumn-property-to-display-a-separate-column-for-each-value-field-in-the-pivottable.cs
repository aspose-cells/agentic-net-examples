// Title: Enable ShowValuesColumn to create separate value columns in an Aspose.Cells PivotTable using C#
// Common Searches: C# code to turn on ShowValuesColumn for an Aspose.Cells PivotTable | How to display each value field in its own column with Aspose.Cells PivotTable | Set PivotTable ShowValuesColumn property via reflection in .NET | Aspose.Cells example for separate value columns in a PivotTable | Enable ShowValuesColumn when refreshing a PivotTable in C# | Programmatic way to show values column in Excel PivotTable using Aspose.Cells
// Tags: Aspose.Cells ShowValuesColumn | C# reflection set PivotTable property | separate value columns in PivotTable | refresh PivotTable Aspose.Cells | programmatic Excel pivot table formatting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an Excel workbook, checks for a PivotTable, uses reflection to set the PivotTable's ShowValuesColumn property to true, refreshes the PivotTable data, and saves the workbook, enabling each value field to appear in its own column.
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Access the first PivotTable
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Enable ShowValuesColumn if the property exists (reflection avoids compile‑time errors)
            var showValuesProp = typeof(PivotTable).GetProperty("ShowValuesColumn");
            if (showValuesProp != null && showValuesProp.CanWrite)
            {
                showValuesProp.SetValue(pivotTable, true);
            }

            // Refresh the PivotTable data using the correct API
            pivotTable.RefreshData();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
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
