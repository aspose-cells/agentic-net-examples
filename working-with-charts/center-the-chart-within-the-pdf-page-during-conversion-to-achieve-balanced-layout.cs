// Title: How to center an Excel chart on a PDF page when converting with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# center chart on PDF page during workbook.Save | Set horizontal and vertical centering for charts in PDF output using Aspose.Cells | Align Excel chart to the middle of a PDF page with Aspose.Cells .NET | Center printed area that contains a chart when saving Excel to PDF in C# | PageSetup.CenterHorizontally not affecting chart placement in PDF Aspose.Cells
// Tags: Aspose.Cells PageSetup centering | C# PDF chart positioning | Excel workbook SaveFormat.Pdf layout | center chart using PageSetup | PDF page layout Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file (or creates a new workbook), accesses the first worksheet, enables both horizontal and vertical centering of the printed area—including any charts—via the PageSetup properties, and then saves the workbook as a PDF, handling Aspose.Cells, I/O, and general exceptions.
class Program
{
    static void Main()
    {
        try
        {
            Workbook workbook;

            // Verify that the input file exists; if not, create a new workbook.
            const string inputPath = "input.xlsx";
            if (File.Exists(inputPath))
            {
                // Load the existing workbook.
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with a default worksheet.
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Access the first worksheet (or the one containing the chart).
            Worksheet worksheet = workbook.Worksheets[0];

            // Center the printed area (including charts) on the PDF page.
            PageSetup pageSetup = worksheet.PageSetup;
            pageSetup.CenterHorizontally = true;
            pageSetup.CenterVertically = true;

            // Save the workbook as a PDF file.
            const string outputPath = "output.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (IOException ex)
        {
            Console.WriteLine($"I/O error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
