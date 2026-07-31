// Title: Aspose.Cells for .NET: Position a chart’s top‑left corner at row 5, column 3 (C#)
// Description: Shows how to build a workbook, add sample data, insert a column chart, and relocate the chart so its upper‑left cell is row 5, column 3 using the Chart.Move method (which updates the Position property) in C# with Aspose.Cells.
// Keywords: Aspose.Cells chart position C# | chart.Move row column Aspose.Cells | set chart top left cell Aspose.Cells | Excel chart placement .NET | Aspose.Cells Position property example
// Common Searches: Aspose.Cells move chart to row 5 column 3 | C# set chart top left corner Aspose.Cells | Chart.Move method Aspose.Cells .NET | How to change chart location programmatically in Excel using Aspose.Cells | Set chart Position property Aspose.Cells C#
// Developer Intent: Move an existing chart so its upper‑left corner aligns with cell C5 (row 5, column 3) on the worksheet.
// Use Cases: Align a chart with a specific report header after data refresh. | Create a tiled dashboard by positioning each chart at predetermined cells. | Enable dynamic chart relocation based on user‑selected rows or columns.
// AI Prompts: Generate C# code that uses Aspose.Cells to set a chart’s Position property so its top‑left corner is at row 5, column 3. | Explain the difference between Chart.Move and the Position property in Aspose.Cells and when each should be used. | Provide an example of moving an existing chart to a new location using cell addresses with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Shows how to build a workbook, add sample data, insert a column chart, and relocate the chart so its upper‑left cell is row 5, column 3 using the Chart.Move method (which updates the Position property) in C# with Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Fruits");
                worksheet.Cells["A3"].PutValue("Vegetables");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(50);
                worksheet.Cells["B3"].PutValue(30);

                // Add a column chart (initial position rows 10‑20, columns 2‑8)
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 10, 2, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set chart data source
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Move the chart so that its top‑left corner is at row 5, column 3
                chart.Move(5, 3, 20, 8);

                // Define output file path
                string outputPath = "ChartTopLeftAtRow5Col3.xlsx";

                // Ensure the output directory exists
                string dir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
