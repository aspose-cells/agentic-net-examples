using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class DataLabelsFromMergedCellsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ---------- Populate source data ----------
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(100);
                worksheet.Cells["B3"].PutValue(200);

                worksheet.Cells["C1"].PutValue("Label");
                worksheet.Cells["C2"].PutValue("100 units");
                worksheet.Cells["C3"].PutValue("200 units");

                // Merge the label cells vertically (C2:C3)
                worksheet.Cells.Merge(1, 2, 2, 1); // rows are zero‑based

                // ---------- Create a column chart ----------
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // ---------- Configure data labels ----------
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;          // show the numeric value
                series.DataLabels.ShowCellRange = true;      // enable pulling from cell range
                series.DataLabels.LinkedSource = "C2:C3";    // range that contains the label text
                series.DataLabels.Font.Color = Color.Blue;  // optional styling

                // ---------- Save the workbook ----------
                string outputPath = "DataLabelsFromMergedCellsDemo.xlsx";
                workbook.Save(outputPath);

                // ---------- Verification ----------
                Cell mergedCell = worksheet.Cells["C2"];
                Console.WriteLine($"Cell C2 IsMerged: {mergedCell.IsMerged}");
                Console.WriteLine($"Merged cell value (should be '100 units'): {mergedCell.StringValue}");
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DataLabelsFromMergedCellsDemo.Run();
        }
    }
}