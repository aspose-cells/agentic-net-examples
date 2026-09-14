// Title: Apply italic formatting to a substring in a chart data label using Aspose.Cells for .NET
// AI Prompts: Generate C# code that builds a column chart, assigns custom text to a data label, and makes the characters inside parentheses italic with Aspose.Cells. | Demonstrate how to use FontSetting.Characters to style a selected range of characters in a chart point's data label and then invoke ApplyFont to update the label. | Create a complete Aspose.Cells example that saves the workbook after applying rich‑text styling to chart data labels.
// Common Searches: how to make part of an Excel chart data label italic with Aspose.Cells C# | Aspose.Cells rich text formatting for chart data labels example | C# apply FontSetting.Characters to a chart point label using Aspose.Cells | save workbook after modifying chart data label rich text Aspose.Cells .NET
// Tags: Aspose.Cells chart data label rich text | apply italic formatting to chart label substring | FontSetting.Characters method Aspose.Cells | DataLabels.ApplyFont after rich text changes | column chart data label custom text C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsRichTextDataLabel
{
    // The sample creates a workbook, adds sample data, inserts a column chart, enables data labels, sets the first label to "120 (est.)", uses FontSetting.Characters to italicize the text inside parentheses, calls DataLabels.ApplyFont to commit the rich‑text changes, and saves the file as RichTextDataLabelItalic.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;

                // Set custom text for the first data label (e.g., "120 (est.)")
                series.Points[0].DataLabels.Text = "120 (est.)";

                // Apply italic formatting to the part "(est.)"
                int start = series.Points[0].DataLabels.Text.IndexOf('(');
                int length = series.Points[0].DataLabels.Text.Length - start;
                if (start >= 0 && length > 0)
                {
                    // Characters returns a FontSetting for the specified range
                    FontSetting labelChars = series.Points[0].DataLabels.Characters(start, length);
                    // Set italic flag for the selected characters
                    labelChars.Font.IsItalic = true;
                }

                // Apply the font changes to all data labels (required after modifying rich text)
                series.DataLabels.ApplyFont();

                // Save the workbook
                string outputPath = "RichTextDataLabelItalic.xlsx";
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
