// Title: C# – Set Aspose.Cells chart top‑left corner to row 5, column 3 using Chart.Position
// Description: Creates a workbook, adds sample data, inserts a column chart, and moves the chart so its top‑left corner aligns with row 5 and column 3 while preserving the original size, then saves the file as an .xlsx document.
// Keywords: Aspose.Cells chart position C# | set chart top left cell Aspose.Cells | Chart.Position Aspose.Cells .NET | move chart without resizing Aspose | Aspose.Cells example chart placement
// Common Searches: Aspose.Cells set chart location row 5 column 3 | C# chart.Position property Aspose.Cells | how to move Aspose.Cells chart to specific cell | Aspose.Cells chart placement example
// Developer Intent: Place a chart’s top‑left corner at row 5, column 3 while keeping its dimensions unchanged.
// Use Cases: Align a sales chart with a header that starts at cell C5 in a financial report. | Build a dashboard where each chart is anchored to precise row/column coordinates. | Re‑position charts after inserting rows so they stay attached to a designated section.
// AI Prompts: Show C# code that uses Chart.Position to set a chart’s top‑left corner to row 5, column 3 in Aspose.Cells. | Give an example of moving an Aspose.Cells chart without changing its size. | Explain the difference between Chart.Move and Chart.Position for positioning charts in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a column chart, and moves the chart so its top‑left corner aligns with row 5 and column 3 while preserving the original size, then saves the file as an .xlsx document.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add some sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Fruits");
                worksheet.Cells["A3"].PutValue("Vegetables");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(50);
                worksheet.Cells["B3"].PutValue(30);

                // Add a column chart. Initial position is rows 10‑20, columns 2‑8
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 10, 2, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Move the chart so that its top‑left corner is at row 5, column 3
                // BottomRow and RightColumn are kept the same as the original size (rows 20, column 8)
                chart.Move(5, 3, 20, 8);

                // Determine output file path
                string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "ChartTopLeftAtRow5Col3.xlsx");

                // Save the workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the chart workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
