// Title: How to make a chart data label italic using Aspose.Cells rich‑text formatting in C#
// Common Searches: Aspose.Cells C# set italic style for chart data label rich text | apply bold or italic formatting to Excel chart data labels with Aspose.Cells | C# code to change font style of column chart data label in Aspose.Cells workbook | using Characters method to format chart data label text in Aspose.Cells | save chart with custom italic data label in .xlsx using Aspose.Cells .NET
// Tags: chart data label italic formatting Aspose.Cells | rich text font setting for Excel chart labels C# | Aspose.Cells Characters method font style | column chart custom data label styling .NET | save workbook with formatted chart labels Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace UpdateRichTextDataLabel
{
    // The example creates a workbook, adds a column chart, enables custom data labels, assigns label text, and uses the Characters method to obtain a FontSetting that makes the label italic and dark blue, then saves the file as RichTextDataLabelItalic.xlsx.
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
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                DataLabels dataLabels = chart.NSeries[0].DataLabels;
                dataLabels.ShowValue = true;               // Show the numeric value
                dataLabels.IsAutoText = false;             // Allow custom text
                dataLabels.Text = "Sample Italic Label";   // Custom label text

                // Apply italic formatting to the entire label text
                // Characters(startIndex, length) returns a FontSetting object
                FontSetting labelFontSetting = dataLabels.Characters(0, dataLabels.Text.Length);
                labelFontSetting.Font.IsItalic = true;
                labelFontSetting.Font.Color = Color.DarkBlue; // optional visual cue

                // Save the workbook
                string outputPath = "RichTextDataLabelItalic.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while updating the rich‑text data label:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
