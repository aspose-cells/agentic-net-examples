// Title: How to insert a column chart into a specific cell range (3 rows × 5 columns) using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that adds a column chart spanning cells C2 to G4 and binds it to the data range A2:B6. | Create a workbook, populate sample data, and place a column chart of exactly three rows by five columns at a defined position using the Aspose.Cells chart API. | Write code to specify the top‑left and bottom‑right cell indices for a chart, set its title, and save the workbook with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# place chart in cells C2 to G4 | set chart dimensions by row and column indices Aspose.Cells .NET | bind chart series to range A2:B6 using Aspose.Cells | create column chart with specific size 3 rows 5 columns in Excel via Aspose.Cells | Aspose.Cells chart placement using firstRow firstColumn parameters
// Tags: Aspose.Cells chart placement by cell range | C# Aspose.Cells column chart size rows columns | Aspose.Cells set chart top left bottom right cells | Aspose.Cells bind chart series to data range | Aspose.Cells save workbook with chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills sample data, inserts a column chart that occupies cells C2 through G4 (three rows by five columns), links the series to the range A2:B6, sets a chart title, and saves the file as ChartInRange.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("Apr");
            sheet.Cells["B5"].PutValue(40);
            sheet.Cells["A6"].PutValue("May");
            sheet.Cells["B6"].PutValue(50);

            // Define chart position: top‑left cell C2 to bottom‑right cell G4
            int firstRow = 1;      // Row index for C2 (zero‑based)
            int firstColumn = 2;   // Column index for C (zero‑based)
            int lastRow = firstRow + 3 - 1;      // Row index for row 4
            int lastColumn = firstColumn + 5 - 1; // Column index for G

            // Add a column chart within the defined range
            int chartIndex = sheet.Charts.Add(ChartType.Column, firstRow, firstColumn, lastRow, lastColumn);
            Chart chart = sheet.Charts[chartIndex];

            // Add series using a range that includes both categories and values (data in columns)
            chart.NSeries.Add("=Sheet1!$A$2:$B$6", false);

            // Optional: set chart title
            chart.Title.Text = "Sample Column Chart";

            // Save the workbook
            string outputPath = "ChartInRange.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
