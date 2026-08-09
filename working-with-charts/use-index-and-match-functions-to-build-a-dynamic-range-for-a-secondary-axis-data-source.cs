// Title: Aspose.Cells .NET: Build a Chart with a Dynamic Secondary Axis Using INDEX/MATCH and Spilled Array
// Description: Shows how to generate an Excel workbook, add a column chart, and use an INDEX/MATCH dynamic array formula to spill the chosen secondary series into a range that is then bound to the chart’s secondary axis at runtime.
// Keywords: Aspose.Cells | .NET | C# | dynamic chart | secondary axis | INDEX MATCH | spilled array formula | chart series range | Excel automation | dynamic range for chart
// Common Searches: Aspose.Cells dynamic secondary axis chart | INDEX MATCH formula for chart series in .NET | spilled array range chart Aspose.Cells | set secondary series values programmatically C# | how to use helper cell for chart data Aspose.Cells
// Developer Intent: Create a column chart where the secondary series data is selected at runtime via an INDEX/MATCH formula and applied as a spilled range.
// Use Cases: End users change a cell value to switch the secondary metric displayed on the chart without code changes. | A template workbook automatically accommodates new data columns for the secondary axis through a single formula. | Interactive dashboards that toggle secondary series on‑the‑fly while keeping primary categories static.
// AI Prompts: Add a dropdown list to cell G1 that lets users pick any secondary series and update the chart automatically. | Explain how to capture the spilled range address after setting the dynamic array formula and reuse it for data labels. | Show how to format the secondary series (color, markers, data labels) while still using the INDEX/MATCH spilled range.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicSecondaryAxis
{
    // Shows how to generate an Excel workbook, add a column chart, and use an INDEX/MATCH dynamic array formula to spill the chosen secondary series into a range that is then bound to the chart’s secondary axis at runtime.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ---------- Populate sample data ----------
                // Primary categories (A2:A5)
                cells["A1"].PutValue("Category");
                cells["A2"].PutValue("Cat1");
                cells["A3"].PutValue("Cat2");
                cells["A4"].PutValue("Cat3");
                cells["A5"].PutValue("Cat4");

                // Primary values (B2:B5)
                cells["B1"].PutValue("Primary");
                cells["B2"].PutValue(10);
                cells["B3"].PutValue(20);
                cells["B4"].PutValue(30);
                cells["B5"].PutValue(40);

                // Secondary series data (D2:F5) with headers in D1:F1
                cells["D1"].PutValue("Sec1");
                cells["E1"].PutValue("Sec2");
                cells["F1"].PutValue("Sec3");

                cells["D2"].PutValue(100);
                cells["D3"].PutValue(200);
                cells["D4"].PutValue(300);
                cells["D5"].PutValue(400);

                cells["E2"].PutValue(500);
                cells["E3"].PutValue(600);
                cells["E4"].PutValue(700);
                cells["E5"].PutValue(800);

                cells["F2"].PutValue(900);
                cells["F3"].PutValue(1000);
                cells["F4"].PutValue(1100);
                cells["F5"].PutValue(1200);

                // Cell G1 will hold the name of the secondary series we want to plot
                cells["G1"].PutValue("Sec2"); // Change this value to pick a different series

                // ---------- Add a chart ----------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Primary series (values from column B)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Add a placeholder for the secondary series
                chart.NSeries.Add("", true);
                chart.NSeries[1].PlotOnSecondAxis = true; // enable secondary axis

                // ---------- Build dynamic range for secondary axis using INDEX/MATCH ----------
                // Formula: =INDEX($D$2:$F$5, , MATCH($G$1, $D$1:$F$1, 0))
                string dynamicFormula = "=INDEX($D$2:$F$5, , MATCH($G$1, $D$1:$F$1, 0))";

                // Place the formula in a helper cell (H2) and let it spill to the required range.
                Cell helperCell = cells["H2"];
                CellArea spillArea = helperCell.SetDynamicArrayFormula(dynamicFormula, new FormulaParseOptions(), true);

                // Construct the spilled range address (e.g., "H2:H5")
                string spilledRange = $"{cells[spillArea.StartRow, spillArea.StartColumn].Name}:{cells[spillArea.EndRow, spillArea.EndColumn].Name}";

                // Use the spilled range as the values for the second series.
                chart.NSeries[1].Values = spilledRange; // set values range for secondary series

                // The secondary series will reuse the primary categories (no need to set CategoryData explicitly).

                // ---------- Save the workbook ----------
                string outputPath = "DynamicSecondaryAxis.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
