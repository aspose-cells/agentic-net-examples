// Title: Aspose.Cells for .NET – Auto‑Resize Chart Data Label Shapes to Fit Updated Text
// Description: C# example that creates a column chart, updates each data label with custom text, and sets IsResizeShapeToFitText so every label shape automatically resizes to fit its content before saving the workbook.
// Keywords: Aspose.Cells | .NET | C# | chart data labels | auto resize data label shape | IsResizeShapeToFitText | ChartPoint | column chart | Excel automation | dynamic label text
// Common Searches: Aspose.Cells auto resize data label shape | IsResizeShapeToFitText example C# | update chart data label text Aspose.Cells | loop through ChartPoint data labels .NET | fit data label shape to text Aspose.Cells
// Developer Intent: Programmatically resize each chart data label shape so it automatically fits the updated label text.
// Use Cases: Generate Excel reports with column charts where label lengths vary and need automatic sizing. | Apply custom prefixes to data label values and ensure labels auto‑adjust without manual formatting. | Iterate over series points to set dynamic label text while keeping the chart visually polished. | Build dashboards that automatically format chart labels for optimal readability.
// AI Prompts: Give C# code using Aspose.Cells to set IsResizeShapeToFitText=true for all data labels in a chart and save the workbook. | Show how to loop through each ChartPoint, assign custom text to DataLabels.Text, and enable auto‑fit of the label shape. | Explain the steps to create a column chart, enable data labels, and make label shapes automatically resize after updating their text with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# example that creates a column chart, updates each data label with custom text, and sets IsResizeShapeToFitText so every label shape automatically resizes to fit its content before saving the workbook.
class ResizeDataLabelsDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(200);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);          // Set values
            chart.NSeries.CategoryData = "A2:A4";      // Set categories

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Update each data label's text and enable auto‑fit to the text
            foreach (ChartPoint point in series.Points)
            {
                // Custom label text
                point.DataLabels.Text = $"Val: {point.YValue}";
                // Allow the label shape to resize automatically to fit its text
                point.DataLabels.IsResizeShapeToFitText = true;
            }

            // Define output file path
            string outputPath = "ResizeDataLabelsDemo.xlsx";

            // Save the workbook with the updated chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        ResizeDataLabelsDemo.Run();
    }
}
