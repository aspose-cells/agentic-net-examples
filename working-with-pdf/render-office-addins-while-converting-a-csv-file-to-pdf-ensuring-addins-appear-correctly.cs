// Title: C# – Convert CSV to PDF with Aspose.Cells while preserving Office add‑ins (charts, images)
// Description: This Aspose.Cells for .NET example loads a CSV file using LoadOptions (LoadFormat.Csv), sets up PdfSaveOptions, and calls ConversionUtility.Convert to generate a PDF that retains any embedded Office add‑ins such as charts, pictures, or shapes.
// Keywords: Aspose.Cells CSV to PDF | preserve add‑ins PDF conversion | PdfSaveOptions charts images | ConversionUtility example .NET | load CSV Aspose.Cells
// Common Searches: Aspose.Cells keep charts when converting CSV to PDF | C# convert CSV to PDF with embedded objects | PdfSaveOptions preserve images Aspose.Cells | How to retain Office add‑ins in PDF output | Sample code CSV to PDF Aspose.Cells .NET
// Developer Intent: Convert a CSV file to PDF in C# and ensure that any Office add‑ins (e.g., charts, pictures, shapes) are rendered in the resulting document.
// Use Cases: Create printable reports from CSV data that include Excel charts. | Batch‑process inventory CSV files into PDF catalogs while keeping product images. | Generate PDF invoices from CSV exports that contain a company logo embedded as an add‑in.
// AI Prompts: Write C# code using Aspose.Cells to convert a CSV file to PDF and verify that embedded charts appear correctly. | Show how to configure PdfSaveOptions to keep images and shapes during CSV‑to‑PDF conversion. | Explain best practices for converting large CSV files with ConversionUtility while preserving all add‑ins.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Saving;

// This Aspose.Cells for .NET example loads a CSV file using LoadOptions (LoadFormat.Csv), sets up PdfSaveOptions, and calls ConversionUtility.Convert to generate a PDF that retains any embedded Office add‑ins such as charts, pictures, or shapes.
class CsvToPdfWithAddIns
{
    static void Main()
    {
        // Input CSV file and output PDF file paths
        string csvPath = "input.csv";
        string pdfPath = "output.pdf";

        // Create a sample CSV file if it does not exist
        if (!System.IO.File.Exists(csvPath))
        {
            System.IO.File.WriteAllText(csvPath,
                "Product,Price,Quantity\nApple,1.20,100\nBanana,0.80,150\nCherry,2.00,75");
        }

        // LoadOptions specify that the source file is a CSV file
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

        // PdfSaveOptions control PDF rendering; add‑ins (charts, images, etc.) are kept by default
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Convert the CSV file to PDF while preserving any embedded objects (add‑ins)
        ConversionUtility.Convert(csvPath, loadOptions, pdfPath, saveOptions);

        Console.WriteLine($"Conversion completed: '{csvPath}' → '{pdfPath}'");
    }
}
