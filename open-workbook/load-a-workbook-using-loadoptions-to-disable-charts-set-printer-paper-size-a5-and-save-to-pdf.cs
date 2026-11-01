// Title: Load an Excel workbook, configure the first worksheet to A5 paper size, and save it as PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to set A5 paper size for a worksheet before converting to PDF | C# load Excel file and export to PDF with custom page setup using Aspose.Cells | example of saving a workbook as PDF with Aspose.Cells while specifying paper size | Aspose.Cells .NET load workbook and change page setup properties programmatically | convert Excel to PDF with A5 page format in C# using Aspose.Cells
// Tags: load workbook Aspose.Cells C# | set worksheet paper size A5 Aspose.Cells | export workbook to PDF Aspose.Cells | custom page setup PDF conversion C# | Aspose.Cells page setup API

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the input Excel file, loads it into an Aspose.Cells Workbook, sets the first worksheet's PageSetup PaperSize to A5, and saves the workbook as a PDF file, handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (LoadDataOnly option not available in this version)
            Workbook workbook = new Workbook(inputPath);

            // Set paper size to A5 for the first worksheet, if any
            if (workbook.Worksheets.Count > 0)
            {
                workbook.Worksheets[0].PageSetup.PaperSize = PaperSizeType.PaperA5;
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
