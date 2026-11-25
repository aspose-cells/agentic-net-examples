// Title: Apply bold font style and center alignment to chart data labels using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set chart data label font to bold | Aspose.Cells center align data label text in a column chart | How to format Excel chart data labels with bold and centered alignment using Aspose.Cells .NET | Apply font size and color to chart series data labels in Aspose.Cells C# | Enable and style data labels for a column chart with Aspose.Cells
// Tags: bold chart data labels Aspose.Cells .NET | center align chart data labels Aspose.Cells | format Excel chart series data labels C# | apply font styling to chart labels Aspose.Cells | Aspose.Cells data label font customization

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelsFormatting
{
    // The example creates a workbook, adds a column chart with sample data, enables data labels for the first series, and formats those labels by applying a bold 12‑point black font and centering the text horizontally before saving the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["A5"].PutValue("D");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Access the first series
                Series series = chart.NSeries[0];

                // Enable data labels
                series.DataLabels.ShowValue = true;

                // Apply bold font style to data labels
                series.DataLabels.Font.IsBold = true;
                // Optional: set a readable font size and color
                series.DataLabels.Font.Size = 12;
                series.DataLabels.Font.Color = Color.Black;

                // Center align the text of data labels
                series.DataLabels.TextHorizontalAlignment = TextAlignmentType.Center;

                // Apply the font changes to all child label nodes
                series.DataLabels.ApplyFont();

                // Save the workbook
                string outputPath = "DataLabelsBoldCentered.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("Data labels are now bold and centered.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
