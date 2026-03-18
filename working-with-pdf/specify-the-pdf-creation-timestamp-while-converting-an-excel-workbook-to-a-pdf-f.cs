using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Create a workbook and add some data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("PDF with custom creation timestamp");

        // Save the workbook to a temporary Excel file (source for conversion)
        string sourcePath = "sample.xlsx";
        workbook.Save(sourcePath, SaveFormat.Xlsx);

        // Load options for the source file (optional, can be omitted)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

        // Prepare PDF save options and set the desired creation time
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            CreatedTime = new DateTime(2023, 1, 1, 12, 0, 0) // example timestamp
        };

        // Destination PDF file path
        string destPath = "output.pdf";

        // Convert Excel to PDF using the specified options
        ConversionUtility.Convert(sourcePath, loadOptions, destPath, pdfOptions);

        Console.WriteLine($"PDF saved with CreatedTime: {pdfOptions.CreatedTime}");
    }
}