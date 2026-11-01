// Title: Generate a worksheet‑by‑worksheet paper size and FitToPages settings report with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# list each worksheet's paper size and FitToPagesWide/Tall values | How to create a page‑setup summary report for all sheets in an Excel file using Aspose.Cells | Export worksheet print settings (PaperSize, FitToPagesWide, FitToPagesTall) to a new workbook in C# | Retrieve and save page setup information for every worksheet with Aspose.Cells .NET | C# generate Excel report of worksheet print layout properties using Aspose.Cells
// Tags: Aspose.Cells page setup extraction | worksheet print settings report C# | export PaperSize FitToPagesWide FitToPagesTall | generate Excel summary workbook Aspose.Cells | auto‑fit columns after data population Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetPaperSizeReport
{
    // The sample loads an existing workbook, iterates through all worksheets, reads each sheet's PageSetup properties (PaperSize, FitToPagesWide, FitToPagesTall), writes the worksheet name and these values into a new workbook, auto‑fits the columns for readability, and saves the result as a report file, handling missing input files and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "WorksheetPaperSizeReport.xlsx";

                // Verify that the source workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(inputPath);

                // Create a new workbook for the report
                Workbook reportWorkbook = new Workbook();

                // Use the first worksheet in the report workbook
                Worksheet reportSheet = reportWorkbook.Worksheets[0];

                // Write header titles
                reportSheet.Cells["A1"].PutValue("Worksheet Name");
                reportSheet.Cells["B1"].PutValue("PaperSize");
                reportSheet.Cells["C1"].PutValue("FitToPagesWide");
                reportSheet.Cells["D1"].PutValue("FitToPagesTall");

                // Start writing data from the second row (zero‑based index)
                int reportRow = 1;

                // Iterate through each worksheet in the source workbook
                foreach (Worksheet ws in sourceWorkbook.Worksheets)
                {
                    PageSetup pageSetup = ws.PageSetup;

                    // Worksheet name
                    reportSheet.Cells[reportRow, 0].PutValue(ws.Name);

                    // Paper size (enum value as string)
                    reportSheet.Cells[reportRow, 1].PutValue(pageSetup.PaperSize.ToString());

                    // FitToPagesWide (-1 means not set)
                    reportSheet.Cells[reportRow, 2].PutValue(pageSetup.FitToPagesWide);

                    // FitToPagesTall (-1 means not set)
                    reportSheet.Cells[reportRow, 3].PutValue(pageSetup.FitToPagesTall);

                    reportRow++;
                }

                // Auto‑fit columns for readability
                reportSheet.AutoFitColumns();

                // Save the report workbook
                reportWorkbook.Save(outputPath);
                Console.WriteLine($"Report saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
