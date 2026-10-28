// Title: Enable row repeat item labels in an Aspose.Cells PivotTable with C# (using reflection for version compatibility)
// Common Searches: C# set RepeatItemLabels property on Aspose.Cells PivotTable rows | Enable repeat item labels for row fields in Aspose.Cells pivot table programmatically | Aspose.Cells pivot table repeat row labels using reflection for older versions | How to make row labels repeat in Excel pivot table with Aspose.Cells API | Set RepeatItemLabels to true in Aspose.Cells C# example
// Tags: Aspose.Cells PivotTable RepeatItemLabels | C# enable row label repetition in pivot table | reflection set PivotTable property Aspose.Cells | Excel workbook pivot table settings Aspose.Cells | Aspose.Cells version compatibility repeat labels

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// This C# example loads an Excel workbook, locates the first PivotTable, and uses reflection to set its RepeatItemLabels property to true when the property is available (Aspose.Cells 22.12+). It handles missing files, absent pivot tables, and version differences, then saves the modified workbook.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook containing the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one pivot table on the worksheet
            if (worksheet.PivotTables.Count > 0)
            {
                // Retrieve the first pivot table
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Attempt to enable repeat item labels if the property exists (Aspose.Cells 22.12+)
                // The property may not be available in older versions; in that case we skip it.
                try
                {
                    // Use reflection to avoid compile‑time dependency on the property
                    var prop = typeof(PivotTable).GetProperty("RepeatItemLabels");
                    if (prop != null && prop.CanWrite)
                    {
                        prop.SetValue(pivotTable, true);
                        Console.WriteLine("RepeatItemLabels property set to true.");
                    }
                    else
                    {
                        Console.WriteLine("RepeatItemLabels property not available in this Aspose.Cells version.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to set RepeatItemLabels: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
