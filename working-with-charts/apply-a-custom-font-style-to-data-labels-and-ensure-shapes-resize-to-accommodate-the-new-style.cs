// Title: Set Custom Font for Chart Data Labels and Auto‑Resize Shapes with Aspose.Cells for .NET
// Description: This example creates a workbook, adds a column chart, enables data labels, applies a Calibri 14 pt bold dark‑blue font, activates automatic shape resizing (IsResizeShapeToFitText), propagates the font to all label nodes (ApplyFont), and saves the file as an XLSX document.
// Keywords: Aspose.Cells chart data label font | IsResizeShapeToFitText | ApplyFont Aspose.Cells | .NET chart label styling | auto resize chart label shape | custom font Excel chart Aspose | Aspose.Cells column chart example
// Common Searches: how to change chart data label font Aspose.Cells .NET | auto resize data label shape Aspose.Cells | set bold Calibri font for Excel chart labels | Aspose.Cells IsResizeShapeToFitText usage | apply font to all data label nodes Aspose
// Developer Intent: Apply a specific font to chart data labels and ensure the label shapes automatically adjust to the new text size using Aspose.Cells for .NET.
// Use Cases: Generate Excel reports with branded chart labels that match corporate typography. | Create dashboards where data labels must remain fully visible after font changes. | Automate workbook creation for clients in the US and Europe who require precise label styling.
// AI Prompts: Show C# code to set a custom font for chart data labels and enable auto‑resize with Aspose.Cells. | Explain the difference between IsResizeShapeToFitText and ApplyFont in Aspose.Cells chart styling. | Provide a step‑by‑step guide to style all data label nodes in a column chart using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a column chart, enables data labels, applies a Calibri 14 pt bold dark‑blue font, activates automatic shape resizing (IsResizeShapeToFitText), propagates the font to all label nodes (ApplyFont), and saves the file as an XLSX document.
    public class CustomDataLabelFontAndResize
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
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

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Apply a custom font style to the data labels
            series.DataLabels.Font.Name = "Calibri";
            series.DataLabels.Font.Size = 14;
            series.DataLabels.Font.Color = Color.DarkBlue;
            series.DataLabels.Font.IsBold = true;

            // Ensure the data label shape resizes to fit the new font
            series.DataLabels.IsResizeShapeToFitText = true;

            // Propagate the font settings to all child label nodes
            series.DataLabels.ApplyFont();

            // Save the workbook
            workbook.Save("CustomDataLabelFontAndResize.xlsx");
        }
    }
}
