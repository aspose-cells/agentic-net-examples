// Title: C# – Load Workbook with LightCells API, Set A3 Paper Size, and Export to PDF using Aspose.Cells
// Description: Demonstrates how to load an existing Excel file with Aspose.Cells LightCells API, configure the default printer and worksheet page‑setup to A3 via LoadOptions, and save the workbook as a PDF document in .NET.
// Keywords: Aspose.Cells | LightCells API | LoadOptions SetPaperSize | PaperA3 | C# Excel to PDF | Excel page setup A3 | .NET PDF export | Workbook.Save PDF | printer paper size | code example
// Common Searches: Aspose.Cells LightCells set paper size A3 | C# load Excel with LightCells and export PDF | how to change printer paper size to A3 in Aspose.Cells | save workbook as PDF with A3 layout .NET | set worksheet page setup paper size programmatically
// Developer Intent: Load an Excel workbook using LightCells, change the printer and worksheet paper size to A3, and generate a PDF file.
// Use Cases: Create A3‑sized PDF reports from Excel templates for high‑resolution printing. | Batch‑convert multiple Excel files to A3 PDF while preserving layout settings. | Generate printable invoices or catalogs in A3 format directly from workbook data without opening Excel.
// AI Prompts: Write C# code that loads an Excel file with LightCells API, sets the default printer paper size to A3, updates each worksheet's PageSetup, and saves the result as a PDF. | Explain how LoadOptions.SetPaperSize affects PDF output when using Aspose.Cells LightCells API. | Show how to make the paper size configurable at runtime based on user input in a C# Aspose.Cells application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Demonstrates how to load an existing Excel file with Aspose.Cells LightCells API, configure the default printer and worksheet page‑setup to A3 via LoadOptions, and save the workbook as a PDF document in .NET.
    public class LightCellsLoadSetPaperSizeAndSavePdf
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Prepare load options and set the default paper size to A3.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.SetPaperSize(PaperSizeType.PaperA3);

            // Load the workbook using LightCells API with the specified load options.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Ensure the workbook's default printer paper size is also set to A3.
            workbook.Settings.PaperSize = PaperSizeType.PaperA3;

            // Set each worksheet's page setup paper size to A3.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
            }

            // Save the workbook as a PDF file.
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF to \"{outputPath}\".");
        }
    }
}
