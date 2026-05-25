using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class CsvToPdfWithCreatedTime
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "input.csv";

        // Path for the output PDF file
        string pdfPath = "output.pdf";

        // Create a sample CSV file (for demonstration purposes)
        System.IO.File.WriteAllText(csvPath,
            "Name,Age,Country\nJohn Doe,30,USA\nJane Smith,25,UK\nBob Lee,40,Canada");

        // Load options to treat the source file as CSV
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

        // PDF save options with a specific historical creation time
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set the creation time to January 1, 2000, 12:00:00 PM
            CreatedTime = new DateTime(2000, 1, 1, 12, 0, 0)
        };

        // Convert the CSV file to PDF using the specified options
        ConversionUtility.Convert(csvPath, loadOptions, pdfPath, pdfOptions);

        Console.WriteLine($"CSV file '{csvPath}' has been converted to PDF '{pdfPath}' with CreatedTime = {pdfOptions.CreatedTime}");
    }
}