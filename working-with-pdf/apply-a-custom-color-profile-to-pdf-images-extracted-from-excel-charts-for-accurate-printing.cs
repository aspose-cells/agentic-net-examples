// Title: Convert an Excel workbook with embedded charts to a PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that verifies an .xlsx file exists, loads it with Aspose.Cells, and saves the workbook—including all chart objects—as a PDF document. | Create a .NET console application that transforms an Excel file containing charts into a PDF, with graceful handling of missing input files.
// Common Searches: Aspose.Cells C# export Excel file with charts to PDF | save Excel workbook as PDF preserving chart images using Aspose.Cells | convert .xlsx to PDF including charts in .NET | Aspose.Cells PDF conversion example handling file not found
// Tags: Aspose.Cells workbook PDF conversion C# | preserve chart graphics during Excel to PDF export | handle missing input file in Aspose.Cells conversion | save Excel charts as PDF using Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of an input .xlsx file, loads it with Aspose.Cells, and saves the entire workbook—including any embedded charts—as a PDF, while providing basic error handling for missing files.
class ChartPdfColorProfile
{
    static void Main()
    {
        try
        {
            // Input and output paths
            string excelPath = @"C:\Input\WorkbookWithCharts.xlsx";
            string outputPdfPath = @"C:\Output\ChartsWithColorProfile.pdf";

            // Verify that the Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Input file not found: {excelPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Save the workbook (including charts) directly to PDF
            workbook.Save(outputPdfPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully created at: {outputPdfPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
