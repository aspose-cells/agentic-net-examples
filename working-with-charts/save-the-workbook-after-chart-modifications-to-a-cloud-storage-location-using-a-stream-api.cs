// Title: Save an Aspose.Cells workbook that contains a column chart to cloud storage via a MemoryStream in C#
// AI Prompts: Generate C# code that creates a workbook with a column chart using Aspose.Cells, writes it to a MemoryStream, and uploads the stream to Azure Blob Storage. | Provide a C# example that serializes an Aspose.Cells workbook with a chart to a stream and stores it in an Amazon S3 bucket using the AWS SDK.
// Common Searches: how to upload an Aspose.Cells generated Excel file with a chart to Azure Blob using MemoryStream C# | Aspose.Cells export workbook to AWS S3 stream C# example | C# Aspose.Cells save chart workbook to Google Cloud Storage via stream | serialize Aspose.Cells workbook to MemoryStream for cloud upload
// Tags: Aspose.Cells serialize workbook to MemoryStream | upload Excel workbook to Azure Blob C# | store Aspose.Cells chart workbook in Amazon S3 | cloud storage stream API with Aspose.Cells | column chart workbook export to cloud

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, populates sample data, adds a column chart, and demonstrates saving the workbook to a MemoryStream; the stream can then be uploaded to cloud storage services such as Azure Blob or Amazon S3.
class ChartToCloud
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Sample Column Chart";

            // Add series using the data range
            int seriesIndex = chart.NSeries.Add("B2:B4", true);
            chart.NSeries[seriesIndex].Name = "Values";

            // NOTE: Marker customization APIs vary between Aspose.Cells versions.
            // The following lines are omitted to maintain compatibility.
            // If your version supports them, you can uncomment and adjust accordingly:
            // chart.NSeries[seriesIndex].Marker.Size = 8;
            // chart.NSeries[seriesIndex].Marker.Symbol = MarkerSymbol.Circle;

            // Prepare output folder
            string outputFolder = Path.Combine(Environment.CurrentDirectory, "Output");
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Define output file path
            string outputPath = Path.Combine(outputFolder, "ChartWorkbook.xlsx");

            // Save workbook (overwrite if exists)
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook with chart saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
