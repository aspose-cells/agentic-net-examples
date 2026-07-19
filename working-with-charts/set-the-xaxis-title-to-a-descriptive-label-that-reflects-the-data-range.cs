// Title: How to Set a Descriptive X‑Axis Title for a Column Chart with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, inserts month‑sales data, adds a column chart, defines the data range, and programmatically sets a visible, formatted X‑axis (CategoryAxis) title with custom text, font, and size before saving the file.
// Keywords: Aspose.Cells X axis title | C# set chart axis label | CategoryAxis.Title Aspose.Cells | Excel column chart axis formatting .NET | programmatic chart axis title | Aspose.Cells chart customization | set axis title from cell | chart axis font Aspose.Cells
// Common Searches: Aspose.Cells set X axis title C# | How to add axis label to Excel chart using Aspose.Cells | CategoryAxis.Title property example | Change font of chart axis title Aspose.Cells .NET | Make chart X axis title visible programmatically
// Developer Intent: Add or modify the X‑axis (category axis) title of an Excel chart programmatically.
// Use Cases: Label months on a sales column chart to clarify the X‑axis range. | Update the X‑axis title dynamically when the displayed data period changes. | Apply corporate font styling to the chart’s X‑axis title for brand consistency.
// AI Prompts: Generate C# code using Aspose.Cells that sets a custom X‑axis title, makes it visible, and applies a specific font and size to a column chart. | Explain how to bind the X‑axis title text to a worksheet cell so the title updates automatically when the cell value changes. | Show how to format the CategoryAxis title with bold, italic, color, and alignment options in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, inserts month‑sales data, adds a column chart, defines the data range, and programmatically sets a visible, formatted X‑axis (CategoryAxis) title with custom text, font, and size before saving the file.
    class SetXAxisTitle
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for the chart
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["B3"].PutValue(1500);
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B4"].PutValue(1800);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories (X‑axis)

                // Set the X‑axis (CategoryAxis) title
                chart.CategoryAxis.Title.Text = "Months (Jan‑Mar)";
                chart.CategoryAxis.Title.IsVisible = true;
                chart.CategoryAxis.Title.Font.Name = "Arial";
                chart.CategoryAxis.Title.Font.Size = 12;

                // Ensure the output directory exists
                string outputPath = "XAxisTitleDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            SetXAxisTitle.Run();
        }
    }
}
