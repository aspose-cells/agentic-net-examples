// Title: C# – Populate an Excel worksheet with category and series data for charts using Aspose.Cells
// Description: This example creates a new workbook, adds a header row (Category, Series1, Series2) and three rows of sample values, ensures the target folder exists, and saves the file as **ChartSourceData.xlsx**. The generated sheet provides ready‑to‑use source data for any chart built with Aspose.Cells.
// Keywords: Aspose.Cells C# chart data | populate Excel source data | write worksheet headers Aspose | save workbook as xlsx | chart source range Aspose.Cells | Excel data for column chart | C# Aspose.Cells example | create chart data programmatically | Excel workbook generation .NET | Aspose.Cells sample code
// Common Searches: how to add category and series rows for a chart with Aspose.Cells | C# example to populate chart source data in Excel | Aspose.Cells write header and values to worksheet | save chart data workbook using Aspose.Cells .NET | populate Excel sheet for chart programmatically
// Developer Intent: Generate an .xlsx file that contains the header row and sample rows needed as the data source for a chart.
// Use Cases: Prepare static data for a column or bar chart by writing categories in column A and two series in columns B and C. | Create a template workbook that downstream services can read to build dynamic charts. | Export test data for validating chart rendering in reporting or dashboard applications.
// AI Prompts: Show how to add additional series columns dynamically to the worksheet for chart data using Aspose.Cells. | Provide code that creates a chart object linked to the range A1:C4 in the same workbook. | Explain how to format the header row (bold, background color) after populating the data.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a new workbook, adds a header row (Category, Series1, Series2) and three rows of sample values, ensures the target folder exists, and saves the file as **ChartSourceData.xlsx**. The generated sheet provides ready‑to‑use source data for any chart built with Aspose.Cells.
    public class PopulateChartData
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate header cells
                worksheet.Cells["A1"].PutValue("Category");   // Category names
                worksheet.Cells["B1"].PutValue("Series1");    // First data series
                worksheet.Cells["C1"].PutValue("Series2");    // Second data series (optional)

                // Populate sample data rows
                worksheet.Cells["A2"].PutValue("Cat1");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["C2"].PutValue(15);

                worksheet.Cells["A3"].PutValue("Cat2");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["C3"].PutValue(25);

                worksheet.Cells["A4"].PutValue("Cat3");
                worksheet.Cells["B4"].PutValue(30);
                worksheet.Cells["C4"].PutValue(35);

                // Define output file path
                string outputPath = "ChartSourceData.xlsx";

                // Ensure the directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            PopulateChartData.Run();
        }
    }
}
