// Title: Create a summary worksheet that lists each sheet’s print area and title rows using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells to add a new worksheet named "PrintAreaReport" that records the PageSetup.PrintArea and PageSetup.PrintTitleRows of every other worksheet in the workbook. | Change the program so that the collected print area and title‑row information is written to a CSV file instead of a new worksheet. | Add robust error handling that logs worksheets with missing print settings and continues processing the rest of the workbook.
// Common Searches: Aspose.Cells how to retrieve print area for each worksheet in C# | C# generate Excel sheet that lists print title rows of all worksheets using Aspose.Cells | list page setup settings of all sheets in a workbook with Aspose.Cells .NET | save worksheet print area report to a new tab using Aspose.Cells | Aspose.Cells .NET get PrintTitleRows property for multiple worksheets
// Tags: worksheet print area extraction Aspose.Cells | print title rows enumeration Aspose.Cells | summary sheet generation page setup Aspose.Cells | CSV export of print settings Aspose.Cells | auto-fit columns report worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPrintAreaReport
{
    // The example loads an existing Excel workbook (or creates a new one), adds a worksheet called "PrintAreaReport", writes column headers, then iterates through all other worksheets to read each sheet's PageSetup.PrintArea and PageSetup.PrintTitleRows. It records these values in the report sheet, auto‑fits the columns for readability, and saves the workbook to the specified output location.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputFile = @"C:\Input\SampleWorkbook.xlsx";
            string outputFile = @"C:\Output\Workbook_With_PrintAreaReport.xlsx";

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputFile);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            Workbook workbook = null;

            try
            {
                // Load the existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputFile))
                {
                    workbook = new Workbook(inputFile);
                }
                else
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    Console.WriteLine("Creating a new workbook with a default worksheet.");
                    workbook = new Workbook(); // creates a workbook with one default sheet
                }

                // Add a new worksheet for the report
                int reportSheetIndex = workbook.Worksheets.Add();
                Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
                reportSheet.Name = "PrintAreaReport";

                // Write header row
                reportSheet.Cells[0, 0].PutValue("Worksheet Name");
                reportSheet.Cells[0, 1].PutValue("Print Area");
                reportSheet.Cells[0, 2].PutValue("Title Rows");

                int reportRow = 1; // Start after header

                // Iterate through all worksheets in the workbook
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Skip the report sheet itself to avoid self‑reference
                    if (ws.Name == reportSheet.Name)
                        continue;

                    // Retrieve the current print area (e.g., "A1:D20")
                    string printArea = ws.PageSetup.PrintArea;
                    if (string.IsNullOrEmpty(printArea))
                        printArea = "Not Set";

                    // Retrieve the title rows configuration (e.g., "$1:$3")
                    string titleRows = ws.PageSetup.PrintTitleRows;
                    if (string.IsNullOrEmpty(titleRows))
                        titleRows = "Not Set";

                    // Populate the report sheet
                    reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                    reportSheet.Cells[reportRow, 1].PutValue(printArea);
                    reportSheet.Cells[reportRow, 2].PutValue(titleRows);

                    reportRow++;
                }

                // Auto‑fit columns for better readability
                reportSheet.AutoFitColumns();

                // Save the workbook with the added report
                workbook.Save(outputFile);
                Console.WriteLine($"Report saved successfully to: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
