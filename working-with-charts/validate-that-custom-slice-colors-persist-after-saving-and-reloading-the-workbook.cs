// Title: Check that custom slice colors in an Aspose.Cells pie chart are retained after saving to XLSX and reloading
// AI Prompts: Create a pie chart, assign distinct colors to each slice via NSeries points, save the workbook as XLSX, reload it, and compare the slice colors to confirm they match. | Replace the pie chart with a doughnut chart, apply custom colors to each segment, perform a save‑and‑load round‑trip, and verify the colors are still preserved. | Add error handling that throws an exception when any reloaded slice color differs from the original custom color, and log the mismatched values.
// Common Searches: Aspose.Cells how to keep pie chart slice colors after saving workbook | C# verify custom colors of chart slices after reloading XLSX with Aspose.Cells | preserve chart segment colors in Aspose.Cells round‑trip save load | test if Aspose.Cells retains custom slice colors in pie chart
// Tags: pie chart slice custom colors Aspose.Cells | XLSX round‑trip chart color persistence | NSeries point color setting Aspose.Cells | verify chart colors after reload C# | Aspose.Cells chart color validation

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Drawing;
using System.IO;

// The example creates a workbook, adds a pie chart with three data points, sets red, green, and blue colors for the slices using NSeries point properties, saves the workbook to an XLSX stream, reloads it, and checks that each slice's foreground color matches the original assignment, reporting whether all custom colors persisted.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            Chart pieChart = sheet.Charts[chartIndex];
            pieChart.NSeries.Add("B2:B4", true);          // Values
            pieChart.NSeries.CategoryData = "A2:A4";      // Categories

            // Define custom colors for each slice
            Color[] customSliceColors = new Color[] { Color.Red, Color.Green, Color.Blue };

            // Apply custom colors to the slices using NSeries
            for (int i = 0; i < customSliceColors.Length; i++)
            {
                pieChart.NSeries[0].Points[i].Area.ForegroundColor = customSliceColors[i];
            }

            // Save the workbook to a memory stream
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.Save(stream, SaveFormat.Xlsx);

                // Reload the workbook from the memory stream
                stream.Position = 0;
                Workbook reloadedWorkbook = new Workbook(stream);
                Worksheet reloadedSheet = reloadedWorkbook.Worksheets[0];
                Chart reloadedChart = reloadedSheet.Charts[chartIndex];

                // Verify that the custom slice colors persisted after reload
                bool colorsMatch = true;
                for (int i = 0; i < customSliceColors.Length; i++)
                {
                    Color loadedColor = reloadedChart.NSeries[0].Points[i].Area.ForegroundColor;
                    if (loadedColor.ToArgb() != customSliceColors[i].ToArgb())
                    {
                        colorsMatch = false;
                        Console.WriteLine($"Slice {i} color mismatch. Expected: {customSliceColors[i]}, Loaded: {loadedColor}");
                    }
                }

                Console.WriteLine(colorsMatch
                    ? "All custom slice colors persisted after saving and reloading."
                    : "Some custom slice colors did not persist.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
