// Title: C# – Hide Rows 20‑25 in an Excel Worksheet and Export to PDF using Aspose.Cells
// Description: Loads an Excel file (creates a simple workbook if missing), hides rows 20‑25 on the first sheet with the zero‑based HideRows method, and saves the result directly as a PDF document via Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# hide rows | Excel to PDF conversion .NET | HideRows method example | zero based row index Aspose | export hidden rows PDF
// Common Searches: Aspose.Cells hide specific rows before PDF export | C# hide rows 20 to 25 in Excel | Convert Excel to PDF after hiding rows | How to use HideRows with Aspose.Cells .NET
// Developer Intent: Remove rows 20‑25 from view in an Excel file and generate a PDF version of the workbook.
// Use Cases: Produce printable reports that omit temporary or draft rows. | Generate clean invoice PDFs where summary rows are hidden. | Automate batch processing to conceal confidential rows before archiving as PDF.
// AI Prompts: Write C# code with Aspose.Cells to hide rows 20‑25 in the first worksheet and save as PDF. | Explain why the HideRows method uses a start index of 19 for row 20. | Add comprehensive error handling for missing input files when converting Excel to PDF after hiding rows.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel file (creates a simple workbook if missing), hides rows 20‑25 on the first sheet with the zero‑based HideRows method, and saves the result directly as a PDF document via Aspose.Cells for .NET.
    public class HideRowsAndSavePdf
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            try
            {
                // Ensure the input file exists; create a simple workbook if missing
                if (!File.Exists(inputPath))
                {
                    var tempWb = new Workbook();
                    tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                    tempWb.Save(inputPath);
                }

                // Load the existing Excel file
                var workbook = new Workbook(inputPath);

                // Access the first worksheet (you can change the index if needed)
                var worksheet = workbook.Worksheets[0];

                // Hide rows 20 to 25 (zero‑based index: start at 19, hide 6 rows)
                worksheet.Cells.HideRows(19, 6);

                // Save the modified workbook as PDF
                workbook.Save(outputPath, SaveFormat.Pdf);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            HideRowsAndSavePdf.Run();
        }
    }
}
