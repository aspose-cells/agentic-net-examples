// Title: Insert a column chart into a specific cell range (C2:G4) spanning three rows and five columns using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a column chart positioned from cell C2 to G4 with Aspose.Cells. | Show how to calculate upper‑left and lower‑right row/column indices so the chart occupies exactly three rows and five columns. | Demonstrate saving the workbook after adding the chart and ensuring the output directory exists.
// Common Searches: Aspose.Cells C# place column chart in range C2:G4 | how to set chart position by row and column indices in Aspose.Cells | define chart area size three rows five columns using Aspose.Cells for .NET | C# Aspose.Cells add chart with custom cell range and save workbook
// Tags: Aspose.Cells add column chart to worksheet | chart positioning using upperLeftRow upperLeftColumn Aspose.Cells | set chart area range C2:G4 Aspose.Cells | save workbook after chart insertion Aspose.Cells | C# Aspose.Cells chart dimensions programmatically

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The program creates a new workbook, fills sample data, calculates the cell coordinates for a chart area covering three rows and five columns (C2:G4), adds a column chart with data from B2:B5, sets a title, ensures the output folder exists, and saves the file as ChartExample.xlsx.
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
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Define chart position: start at row 1, column 2 (zero‑based) and span 3 rows × 5 columns
            int upperLeftRow = 1;          // Excel row 2
            int upperLeftColumn = 2;       // Excel column C
            int lowerRightRow = upperLeftRow + 2;   // three rows total
            int lowerRightColumn = upperLeftColumn + 4; // five columns total

            // Add a column chart within the defined range
            int chartIndex = sheet.Charts.Add(ChartType.Column, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart series (values from B2:B5, categories from A2:A5)
            int seriesIndex = chart.NSeries.Add("B2:B5", true);
            // Note: In recent Aspose.Cells versions, CategoryData is set automatically when the range includes categories.
            // If needed, you can set it via chart.NSeries[seriesIndex].CategoryData = "A2:A5";

            // Optional: set a title for the chart
            chart.Title.Text = "Sample Column Chart";

            // Determine output file path
            string outputPath = "ChartExample.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
