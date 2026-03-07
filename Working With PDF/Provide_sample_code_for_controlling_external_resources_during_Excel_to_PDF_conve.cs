using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class ExcelToPdfConversionWithExternalResources
{
    public static void Run()
    {
        // Paths for the temporary CSV source and the resulting PDF
        string sourceCsv = "sample.csv";
        string destPdf = "output.pdf";

        // Create a simple CSV file that will be used as the source document
        File.WriteAllText(sourceCsv,
            "Name,Description\n" +
            "Item1,\"Description with a link http://example.com\"\n" +
            "Item2,\"Another description\"");

        // LoadOptions tell Aspose.Cells how to interpret the source file (CSV format)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

        // Configure PDF save options to control handling of external resources
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed any OLE attachments (if the source contained them) into the PDF
            EmbedAttachments = true,

            // Suppress rendering errors (e.g., missing images) so conversion continues
            IgnoreError = true
        };

        // Perform the conversion using the provided ConversionUtility method
        ConversionUtility.Convert(sourceCsv, loadOptions, destPdf, pdfOptions);

        Console.WriteLine($"Conversion completed successfully. PDF saved to '{destPdf}'.");

        // Clean up the temporary CSV file
        if (File.Exists(sourceCsv))
        {
            File.Delete(sourceCsv);
        }
    }

    static void Main(string[] args)
    {
        Run();
    }
}