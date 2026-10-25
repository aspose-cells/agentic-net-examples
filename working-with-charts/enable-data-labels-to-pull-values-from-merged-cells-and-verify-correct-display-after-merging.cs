// Title: Display merged cell category as data label in an Aspose.Cells column chart (C#)
// Common Searches: C# Aspose.Cells chart using merged cells for category labels | How to show merged cell text as data label in an Aspose.Cells column chart | Enable ShowCategoryName for a merged range in Aspose.Cells C# example | Create Excel column chart with merged category cells using Aspose.Cells | Aspose.Cells data labels display merged cell values in C#
// Tags: Aspose.Cells merged cells data labels | column chart category from merged range | C# Aspose.Cells NSeries ShowCategoryName | Excel chart merged cell category label | Aspose.Cells chart series with merged cells | C# create column chart from merged cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, merges cells A2:A4 to form a single category label, adds a column chart that uses the merged range for categories and B2:B4 for values, enables data labels to display the merged category name while hiding numeric values, and saves the result as MergedCellDataLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook(); // create empty workbook
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Populate data ----------
            // Header
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            // Category values (will be merged)
            sheet.Cells["A2"].PutValue("Merged Category");
            sheet.Cells["A3"].PutValue(string.Empty); // empty cells that will be merged
            sheet.Cells["A4"].PutValue(string.Empty);

            // Numeric values
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // ---------- Merge cells for the category label ----------
            // Merge A2:A4 so the category label comes from a single merged cell
            sheet.Cells.Merge(1, 0, 3, 0); // rows 1‑3, column 0 (A2:A4)

            // ---------- Create a column chart ----------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add series with both category and value ranges (A2:A4 for categories, B2:B4 for values)
            chart.NSeries.Add("A2:A4,B2:B4", false);

            // ---------- Enable data labels to show category names ----------
            chart.NSeries[0].DataLabels.ShowCategoryName = true; // show merged category as label
            chart.NSeries[0].DataLabels.ShowValue = false;       // hide numeric value if not needed

            // ---------- Save the workbook ----------
            string outputPath = "MergedCellDataLabels.xlsx";
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
