using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class AddSignatureLineToChartSheet
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate the first worksheet with data for the chart
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["A4"].PutValue("C");
            dataSheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = dataSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = dataSheet.Charts[chartIndex];
            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Create and configure a signature line
            SignatureLine signatureLine = new SignatureLine
            {
                Signer = "John Doe",
                Title = "Approver",
                Email = "john.doe@example.com",
                IsLine = true,
                AllowComments = true,
                ShowSignedDate = true,
                Instructions = "Please sign to approve the chart."
            };

            // Add the signature line to the worksheet at a specific cell position (row 5, column 5)
            Picture signaturePicture = dataSheet.Shapes.AddSignatureLine(5, 5, signatureLine);
            // Optionally set the size of the signature picture
            signaturePicture.Width = 150;
            signaturePicture.Height = 50;

            // Save the workbook
            string outputPath = "ChartSheetWithSignatureLine.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}