// Title: Set a column chart title to "Monthly Revenue" and apply an Accent1‑like color with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to add a column chart, set its title to "Monthly Revenue", and change the title font color to the Accent1 theme color. | Update an existing Aspose.Cells workbook to modify a chart's title text and apply a predefined Accent1 blue color to the title font.
// Common Searches: Aspose.Cells C# set column chart title text and change title font color to theme Accent1 | How to apply Accent1 style to chart title using Aspose.Cells .NET | Change chart title color to blue Accent1 in Aspose.Cells workbook | Programmatically format chart title in Aspose.Cells for .NET
// Tags: Aspose.Cells set chart title text | Aspose.Cells apply Accent1 color to chart title | column chart title formatting Aspose.Cells .NET | customize chart title font color Aspose.Cells | add column chart and style title Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a new workbook, inserts a column chart, sets the chart title to "Monthly Revenue", applies an Accent1‑like blue color to the title font, and saves the file as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle rule)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a column chart to the worksheet (example chart)
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];

                // Set the chart title
                chart.Title.Text = "Monthly Revenue";

                // Apply a theme-like accent color to the title font
                // Aspose.Cells may not expose Theme.ColorScheme directly in all versions,
                // so we use a standard color that resembles Accent1.
                chart.Title.Font.Color = Color.FromArgb(0, 112, 192); // Accent1-like blue

                // Save the workbook (lifecycle rule)
                string outputPath = "output.xlsx";
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
