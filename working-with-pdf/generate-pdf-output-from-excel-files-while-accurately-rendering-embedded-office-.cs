using System;
using System.IO;
using Aspose.Cells;

class ExcelToPdfWithAddIns
{
    static void Main()
    {
        // Path to the Excel file that contains embedded Office Add‑Ins
        string sourceFile = Path.Combine(Environment.CurrentDirectory, "input_with_addins.xlsx");

        // Path where the resulting PDF will be saved
        string outputFile = Path.Combine(Environment.CurrentDirectory, "output.pdf");

        if (!File.Exists(sourceFile))
        {
            Console.WriteLine($"Source file not found: {sourceFile}");
            return;
        }

        // Load the workbook
        Workbook workbook = new Workbook(sourceFile);

        // Set PDF save options (default options preserve embedded objects as images)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save as PDF
        workbook.Save(outputFile, pdfOptions);

        Console.WriteLine($"Conversion completed successfully. PDF saved to '{outputFile}'.");
    }
}