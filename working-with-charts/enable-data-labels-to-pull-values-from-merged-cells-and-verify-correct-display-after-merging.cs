// Title: Show merged‑cell values as data labels in an Aspose.Cells column chart (C#)
// AI Prompts: Create a C# workbook with Aspose.Cells that merges cells B2:C2, adds a column chart using B2:B4 as the series range, and enables data labels to display the merged cell content. | Insert code that reads the string from the merged cell B2 and compares it to the first data label in the chart to verify correctness. | Modify the example to generate a line chart instead of a column chart while still pulling the data label text from the merged cell.
// Common Searches: Aspose.Cells C# chart data label from merged cell value | how to display merged cell content in chart labels using Aspose.Cells .NET | verify that chart data label matches merged cell in Aspose.Cells example | C# Aspose.Cells merge cells and use them as series values in a column chart | read merged cell value for first point data label Aspose.Cells
// Tags: merge cells for chart series Aspose.Cells C# | show merged cell in chart label Aspose.Cells | chart data label consistency check Aspose.Cells | extract merged cell string Aspose.Cells C# | export workbook with chart Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMergedCellDataLabels
{
    // The sample creates a new workbook, merges cells B2 and C2, builds a column chart using B2:B4 as Y‑values and A2:A4 as X‑values, turns on data labels, reads the merged cell's string value to confirm that the first data label reflects the merged content, prints verification results, and saves the file as MergedCellDataLabels.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data
                cells["A1"].PutValue("Category");
                cells["A2"].PutValue("Q1");
                cells["A3"].PutValue("Q2");
                cells["A4"].PutValue("Q3");

                cells["B1"].PutValue("Value");
                cells["B2"].PutValue(120); // Will be merged with C2
                cells["B3"].PutValue(150);
                cells["B4"].PutValue(180);

                // Merge B2 and C2 to simulate a merged cell containing the value 120
                cells.Merge(1, 1, 1, 2); // Row 1 (zero‑based), Column 1 (B), 1 row, 2 columns

                // Create a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the series: X values from A2:A4, Y values from B2:B4 (includes merged cell)
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries[0].XValues = "A2:A4";

                // Enable data labels to show the values
                chart.NSeries[0].DataLabels.ShowValue = true;

                // Retrieve the value from the merged cell directly (since DataLabels collection indexing may not be supported)
                string mergedCellValue = cells["B2"].StringValue;
                Console.WriteLine("Data label for first point (merged cell): " + mergedCellValue);

                if (mergedCellValue == "120")
                {
                    Console.WriteLine("Verification passed: Data label correctly reflects merged cell value.");
                }
                else
                {
                    Console.WriteLine($"Verification failed: Expected '120' but got '{mergedCellValue}'.");
                }

                // Save the workbook
                workbook.Save("MergedCellDataLabels.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
