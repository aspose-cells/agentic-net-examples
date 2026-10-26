// Title: C# example: load an Excel workbook with Aspose.Cells, set all worksheets to A3 paper size, and export to PDF
// Common Searches: how to change page setup to A3 for every sheet before converting Excel to PDF using Aspose.Cells in C# | Aspose.Cells C# set printer paper size A3 when saving workbook as PDF | load Excel file and adjust paper size for PDF output with Aspose.Cells .NET | C# code to iterate worksheets and set PaperSizeType.PaperA3 before PDF export | export Excel to PDF with custom page size using Aspose.Cells for .NET
// Tags: Aspose.Cells set worksheet paper size A3 | C# export Excel to PDF with custom page setup | Aspose.Cells page setup before PDF conversion | adjust printer settings Aspose.Cells .NET | save workbook as PDF with specific paper size

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, iterates through each worksheet to set the PageSetup.PaperSize to PaperSizeType.PaperA3, and then saves the workbook as a PDF file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (standard loading; LightCells not required for this example)
            var workbook = new Workbook(inputPath);

            // Set printer paper size to A3 for all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
