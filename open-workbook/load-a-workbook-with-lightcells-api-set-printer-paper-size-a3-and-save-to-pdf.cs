// Title: Load Workbook with LightCells, Set A3 Paper Size, and Export to PDF – Aspose.Cells for .NET
// Description: Demonstrates how to use Aspose.Cells LightCells API in C# to load an Excel file, configure LoadOptions to PaperA3, set the workbook and each worksheet to A3 printer paper size, and save the result as a PDF. Ideal for fast processing of large workbooks with custom page layouts.
// Keywords: Aspose.Cells | LightCells API | LoadOptions.SetPaperSize | PaperA3 | C# | .NET | Excel to PDF conversion | printer paper size | worksheet PageSetup | fast workbook loading | PDF export
// Common Searches: How to load an Excel file with LightCells and change paper size to A3 before PDF conversion | Aspose.Cells set default printer paper size A3 for all worksheets | Convert Excel to PDF with A3 page size using Aspose.Cells .NET | LightCells LoadOptions paper size example
// Developer Intent: Load an Excel workbook via LightCells, set its printer paper size to A3, and save it as a PDF.
// Use Cases: Generate printable A3 PDF reports from large Excel workbooks with minimal memory footprint. | Apply a uniform A3 page layout to every worksheet before batch PDF conversion in a server‑side application. | Ensure consistent printer settings when converting user‑uploaded Excel files to PDF in a web service.
// AI Prompts: Provide C# code that uses Aspose.Cells LightCells API to open an Excel file, set the default and worksheet page setup to PaperA3, and save as PDF. | Explain how LoadOptions.SetPaperSize affects printing settings when a workbook is loaded with LightCells. | Show how to modify the example to use a different paper size, such as Letter, while keeping LightCells performance benefits.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Demonstrates how to use Aspose.Cells LightCells API in C# to load an Excel file, configure LoadOptions to PaperA3, set the workbook and each worksheet to A3 printer paper size, and save the result as a PDF. Ideal for fast processing of large workbooks with custom page layouts.
    public class LightCellsLoadSetPaperSizeAndSavePdf
    {
        public static void Run()
        {
            try
            {
                // Path to the source Excel file
                string inputPath = "input.xlsx";

                // Ensure the input file exists; create a simple workbook if it does not.
                if (!File.Exists(inputPath))
                {
                    var tempWb = new Workbook();
                    tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                    tempWb.Save(inputPath);
                }

                // Create LoadOptions and set the default paper size to A3.
                // This influences the workbook's default print settings when loaded via LightCells API.
                LoadOptions loadOptions = new LoadOptions();
                loadOptions.SetPaperSize(PaperSizeType.PaperA3);

                // Load the workbook using the LightCells API (constructor that accepts LoadOptions).
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Ensure the workbook's default printer paper size is set to A3.
                workbook.Settings.PaperSize = PaperSizeType.PaperA3;

                // Optionally, set each worksheet's page setup to A3 as well.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
                }

                // Save the workbook as a PDF file.
                string outputPath = "output.pdf";
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine($"Workbook loaded with LightCells, paper size set to A3, and saved to PDF at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            LightCellsLoadSetPaperSizeAndSavePdf.Run();
        }
    }
}
