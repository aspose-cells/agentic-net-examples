// Title: Convert an Excel workbook to PDF with Aspose.Cells for .NET (C#) and handle missing file errors
// Common Searches: c# convert xlsx to pdf using Aspose.Cells with file existence check | aspocells workbook.save pdf example with error handling in .NET | how to verify Excel file before converting to PDF with Aspose.Cells | Aspose.Cells SaveFormat.Pdf usage in C# application | generate PDF from Excel workbook using Aspose.Cells and catch exceptions
// Tags: Aspose.Cells PDF export operation | C# Excel to PDF conversion using Aspose.Cells | Validate input file before Aspose.Cells conversion | Exception handling for Aspose.Cells PDF generation | SaveFormat.Pdf parameter in Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the source Excel file exists, loads it into an Aspose.Cells Workbook, and saves the workbook as a PDF using SaveFormat.Pdf, while providing basic exception handling for any runtime errors.
class ExcelToPdfWithAttachment
{
    static void Main()
    {
        try
        {
            const string excelPath = "input.xlsx";
            const string outputPdfPath = "output.pdf";

            // Verify that the source Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(excelPath);

            // Convert and save the workbook as PDF
            workbook.Save(outputPdfPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully created at \"{outputPdfPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
