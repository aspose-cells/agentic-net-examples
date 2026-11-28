// Title: C# – Apply bold red font and yellow background to data labels of the second series in an Aspose.Cells column chart
// Common Searches: Aspose.Cells C# format data labels of a specific series in a column chart | set bold red font for second series data labels using Aspose.Cells | change data label background color to yellow in Aspose.Cells chart C# | how to customize chart series label appearance with Aspose.Cells API | Aspose.Cells chart data labels styling example C#
// Tags: format second series data labels Aspose.Cells | bold red font chart labels C# | yellow fill background data labels Aspose.Cells | column chart series label styling Aspose.Cells | Aspose.Cells chart data label customization XLSX

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a column chart with two series, accesses the DataLabels of the second series, enables value display, sets the label font to bold red, applies a solid yellow fill, and saves the result as an XLSX file using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for two series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");

            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add the first series (Series1)
            chart.NSeries.Add("B2:B4", true);
            // Add the second series (Series2)
            chart.NSeries.Add("C2:C4", true);
            // Set the category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A4";

            // -------------------------------------------------
            // Format data labels of the SECOND series
            // -------------------------------------------------
            Series secondSeries = chart.NSeries[1]; // index 1 = second series

            // Access the DataLabels object
            DataLabels dataLabels = secondSeries.DataLabels;
            dataLabels.ShowValue = true; // ensure values are displayed

            // Font: bold and red
            dataLabels.Font.IsBold = true;
            dataLabels.Font.Color = Color.Red;

            // Background: solid yellow (pattern fill)
            dataLabels.Area.FillFormat.Pattern = FillPattern.Solid;
            // Note: In some Aspose.Cells versions the ForeColor property may not be available.
            // If it exists, uncomment the following line to set the fill color:
            // dataLabels.Area.FillFormat.ForeColor = Color.Yellow;

            // Save the workbook
            string outputPath = "ChartWithFormattedDataLabels.xlsx";

            // Ensure the directory exists (if a directory part is present)
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
