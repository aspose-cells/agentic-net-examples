// Title: Position an Excel chart at row 15, column 3 and set its width to 400 points with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# place chart at row 15 column 3 | set Excel chart width in points using Aspose.Cells .NET | move chart to a specific cell index with Aspose.Cells | adjust chart size and location programmatically in .NET | UpperLeftRow UpperLeftColumn properties example for Aspose.Cells
// Tags: Aspose.Cells chart positioning by cell index | set chart width points Aspose.Cells .NET | place Excel chart at specific row column Aspose.Cells | chart object size adjustment Aspose.Cells | UpperLeftRow UpperLeftColumn Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartPositionExample
{
    // The example creates a workbook, adds sample data, inserts a column chart, then uses the ChartObject's UpperLeftRow and UpperLeftColumn properties to position the chart at row 15, column 3 (zero‑based indices) and sets its Width to 400 points (Height to 300 points) before saving the file as ChartPositioned.xlsx.
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

                // Add a column chart (initial position is temporary)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 0, 0, 10, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Position the chart at row 15, column 3
                // Use UpperLeftRow and UpperLeftColumn to set the top‑left cell of the chart
                chart.ChartObject.UpperLeftRow = 15;   // zero‑based index (row 16 in UI)
                chart.ChartObject.UpperLeftColumn = 3; // zero‑based index (column D)

                // Set the size of the chart (width in points, height in points)
                chart.ChartObject.Width = 400;
                chart.ChartObject.Height = 300;

                // Save the workbook
                string outputPath = "ChartPositioned.xlsx";

                // Ensure the directory exists (prevent FileNotFoundException)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
