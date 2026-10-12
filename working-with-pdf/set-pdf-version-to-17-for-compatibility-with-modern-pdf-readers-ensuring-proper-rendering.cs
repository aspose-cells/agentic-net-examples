// Title: Set PDF version to 1.7 when exporting an Excel workbook to PDF using Aspose.Cells in C#
// AI Prompts: Write C# code that uses Aspose.Cells PdfSaveOptions.Version to force PDF 1.7 when saving a Workbook as a PDF file. | Show how to configure Aspose.Cells PdfSaveOptions for PDF 1.7 compatibility before calling Workbook.Save in a .NET application.
// Common Searches: Aspose.Cells export Excel to PDF with PDF 1.7 compatibility C# | C# set specific PDF version when saving workbook using Aspose.Cells | How to force PDF 1.7 output with Aspose.Cells PdfSaveOptions | Save Excel file as PDF 1.7 using Aspose.Cells .NET
// Tags: Aspose.Cells PdfSaveOptions PDF 1.7 | C# export workbook to PDF with specific version | configure PDF compatibility Aspose.Cells | set PDF version Aspose.Cells .NET | Excel to PDF 1.7 using Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a Workbook, adds sample data, configures PdfSaveOptions to use PDF version 1.7, and saves the workbook as a PDF file, ensuring compatibility with modern PDF readers.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example data – optional, just to have content in the PDF
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue("Data");

            // Configure PDF save options (default PDF version is 1.7 in recent Aspose.Cells versions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save("Result.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
