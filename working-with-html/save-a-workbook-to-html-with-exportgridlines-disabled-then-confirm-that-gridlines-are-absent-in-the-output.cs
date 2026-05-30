using System;
using System.IO;
using Aspose.Cells;

class ExportGridLinesDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Make gridlines visible in the worksheet (they would be exported if enabled)
        worksheet.IsGridlinesVisible = true;

        // Add some sample data
        worksheet.Cells["A1"].PutValue("Sample Data");
        worksheet.Cells["B2"].PutValue(123);

        // Configure HTML save options with ExportGridLines disabled
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportGridLines = false // Ensure gridlines are not exported
        };

        string htmlFilePath = "output.html";

        // Save the workbook as HTML using the configured options
        workbook.Save(htmlFilePath, htmlOptions);

        // Read the generated HTML file
        string htmlContent = File.ReadAllText(htmlFilePath);

        // Simple verification: check that typical gridline CSS is not present
        bool gridLinesFound = htmlContent.Contains("border:") && htmlContent.Contains("#d0d0d0");
        Console.WriteLine("Gridlines exported: " + gridLinesFound);
        Console.WriteLine("ExportGridLines disabled, gridlines absent: " + (!gridLinesFound));
        Console.WriteLine("HTML file saved at: " + Path.GetFullPath(htmlFilePath));
    }
}