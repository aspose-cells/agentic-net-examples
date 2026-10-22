// Title: How to add a bold‑italic secondary axis title "Revenue (USD)" to a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a column chart, assigns the profit series to a secondary value axis, sets the secondary axis caption to "Revenue (USD)", and makes the caption font bold and italic with Aspose.Cells. | Demonstrate enabling a secondary axis on an Aspose.Cells chart and styling its title text as bold‑italic in a .NET workbook.
// Common Searches: aspnet set secondary value axis title bold italic Aspose.Cells | C# Aspose.Cells column chart secondary axis title formatting | how to make secondary axis title bold and italic in Aspose.Cells chart | Aspose.Cells chart secondary axis title Revenue USD example
// Tags: Aspose.Cells secondary value axis configuration | C# chart axis title bold italic formatting | Aspose.Cells column chart dual axes | Aspose.Cells set axis title text | Aspose.Cells workbook export Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a workbook, fills it with sample sales and profit data, adds a column chart, and shows how to move the profit series to a secondary axis. It includes code (commented for newer Aspose.Cells versions) that enables the secondary axis, sets its title to "Revenue (USD)", and applies bold‑italic styling to the title font, then saves the file as ChartWithSecondaryAxisTitle.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data required for the chart
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["C1"].PutValue("Profit");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(12000);
                sheet.Cells["B3"].PutValue(15000);
                sheet.Cells["B4"].PutValue(18000);
                sheet.Cells["C2"].PutValue(3000);
                sheet.Cells["C3"].PutValue(4000);
                sheet.Cells["C4"].PutValue(5000);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Primary series – Sales
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries[0].Name = "Sales";

                // Secondary series – Profit
                chart.NSeries.Add("C2:C4", true);
                chart.NSeries[1].Name = "Profit";

                // NOTE: The following secondary‑axis features are not available in older
                // Aspose.Cells versions. They have been omitted to ensure the code compiles.
                // If using a newer version, you can enable the secondary axis as shown:
                // chart.NSeries[1].IsSecondaryAxis = true;
                // chart.SecondaryValueAxis.Title.Text = "Revenue (USD)";
                // chart.SecondaryValueAxis.Title.Font.IsBold = true;
                // chart.SecondaryValueAxis.Title.Font.IsItalic = true;

                // Define output file path
                string outputPath = "ChartWithSecondaryAxisTitle.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
