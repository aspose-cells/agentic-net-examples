using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class XlsToPdfConverter
{
    static void Main()
    {
        // Path to the source XLS file that contains images and/or charts
        string sourcePath = "input.xls";

        // Path where the resulting PDF will be saved
        string destPath = "output.pdf";

        // Simple conversion using the two‑parameter overload.
        ConversionUtility.Convert(sourcePath, destPath);
        Console.WriteLine("Simple conversion completed: " + destPath);

        // Conversion with explicit load and save options.
        // LoadOptions without specifying format will auto‑detect the source file type.
        LoadOptions loadOptions = new LoadOptions();

        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true,
            EmbedAttachments = true
        };

        // Perform the conversion with the specified options.
        ConversionUtility.Convert(sourcePath, loadOptions, destPath, pdfOptions);
        Console.WriteLine("Conversion with options completed: " + destPath);
    }
}