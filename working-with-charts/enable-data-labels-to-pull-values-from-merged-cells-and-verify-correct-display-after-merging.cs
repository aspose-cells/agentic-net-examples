// Title: Show a merged cell value as the category label for all points in an Aspose.Cells column chart (C#)
// Common Searches: how to use a merged cell as XValues for a column chart in Aspose.Cells C# | Aspose.Cells display the same category name for every data point after merging cells | C# chart series data labels showing merged cell text instead of individual cell values | set chart series XValues to B2:B5 merged range Aspose.Cells | verify merged cell label appears correctly in Aspose.Cells generated Excel file
// Tags: assign XValues to merged cells Aspose.Cells C# | show category name on data labels from merged cell | column chart series with unified merged label | check merged cell label rendering in Excel chart | Aspose.Cells handling of merged cells for chart labels

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMergedLabelDemo
{
    // The example creates a workbook, fills column A with numeric data and column B with a label that is merged across rows B2:B5, then adds a column chart. The series uses A2:A5 for values and sets XValues to the merged cells B2:B5. Data labels are configured to display the category name (the merged label) while hiding the numeric values. The expected label is printed for verification before saving the workbook as MergedLabelChart.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (create rule)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate data values in column A
                cells["A1"].PutValue("Value");
                cells["A2"].PutValue(10);
                cells["A3"].PutValue(20);
                cells["A4"].PutValue(30);
                cells["A5"].PutValue(40);

                // Populate label text in column B (to be merged)
                cells["B1"].PutValue("Label");
                cells["B2"].PutValue("MergedLabel"); // This value will be used for all data points

                // Merge cells B2:B5 so that the label is a single merged cell
                sheet.Cells.Merge(1, 1, 4, 1); // rows 1‑4 (zero‑based), column 1 (B)

                // Create a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add a series to the chart (values from A2:A5)
                int seriesIndex = chart.NSeries.Add("A2:A5", true);
                Series series = chart.NSeries[seriesIndex];

                // Set the X‑values (category labels) to the merged range B2:B5
                series.XValues = "B2:B5";

                // Enable data labels to show the category (label) name
                series.DataLabels.ShowCategoryName = true;
                series.DataLabels.ShowValue = false; // hide the numeric value for clarity

                // Verify that the data label text matches the merged cell value
                // After merging, all cells in B2:B5 return the value of the top‑left cell (B2)
                string expectedLabel = cells["B2"].StringValue;
                Console.WriteLine("Expected data label (merged cell value): " + expectedLabel);

                // Save the workbook (save rule)
                string outputPath = "MergedLabelChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
