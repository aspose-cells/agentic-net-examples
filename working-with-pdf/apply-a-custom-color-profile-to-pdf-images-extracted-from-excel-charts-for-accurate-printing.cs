// Title: Export Excel charts to PDF with Aspose.Cells for .NET (C#)
// Common Searches: aspnet c# export excel chart as pdf using Aspose.Cells | save workbook with embedded charts to pdf Aspose.Cells example | convert Excel file containing charts to PDF in C# | Aspose.Cells generate PDF from Excel workbook including chart images | C# code to create PDF from Excel charts with Aspose.Cells
// Tags: excel chart pdf export aspose.cells c# | save workbook as pdf with charts c# | aspose.cells generate pdf from excel charts | c# pdf conversion of excel workbook including charts | aspose.cells saveformat.pdf usage example

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, verifies its existence, and saves the entire workbook—including any embedded charts—as a PDF file using Aspose.Cells' SaveFormat.Pdf in C#, with basic error handling.
class ExcelChartToPdfWithColorProfile
{
    static void Main()
    {
        try
        {
            // Paths to required files
            string excelPath = "ChartData.xlsx";
            string pdfOutputPath = "ChartsWithColorProfile.pdf";

            // Verify that the source Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Save the workbook (including charts) as PDF
            workbook.Save(pdfOutputPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully created at '{pdfOutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
