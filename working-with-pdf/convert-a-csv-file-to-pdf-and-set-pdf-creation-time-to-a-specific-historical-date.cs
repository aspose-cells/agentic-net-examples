using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Author: Example code for CSV to PDF conversion with custom CreatedTime

        // Path to the source CSV file
        string csvFilePath = "input.csv";

        // Load the CSV file into a Workbook instance
        // LoadOptions with LoadFormat.Csv ensures proper CSV parsing
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
        Workbook workbook = new Workbook(csvFilePath, loadOptions);

        // Configure PDF save options and set a specific historical creation time
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example historical date: May 23, 1995, 10:30:00 AM
            CreatedTime = new DateTime(1995, 5, 23, 10, 30, 0)
        };

        // Path for the output PDF file
        string pdfFilePath = "output.pdf";

        // Save the workbook as PDF using the configured options
        workbook.Save(pdfFilePath, pdfOptions);
    }
}