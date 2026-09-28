// Title: Generate a C# Aspose.Cells report of each worksheet’s paper size, FitToPagesWide, and FitToPagesTall settings
// AI Prompts: Write C# code with Aspose.Cells that opens an existing workbook, reads each worksheet’s PageSetup object, and creates a new workbook listing the worksheet name, page size setting, FitToPagesWide, and FitToPagesTall. | Extend the program to also capture the PageSetup orientation and margin values and add them to the report workbook. | Modify the script to skip worksheets where FitToPagesWide is set to 1 and only include sheets with custom scaling in the output file.
// Common Searches: how to extract worksheet page setup properties using Aspose.Cells in C# | C# Aspose.Cells list paper size and fit-to-page settings for all sheets | generate Excel file with page setup summary for each worksheet using Aspose.Cells | retrieve FitToPagesWide and FitToPagesTall values from a workbook with Aspose.Cells .NET
// Tags: Aspose.Cells extract worksheet PageSetup data | C# create page size summary workbook | Aspose.Cells read FitToPagesWide property | export scaling settings to new Excel file | list worksheet scaling values Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program loads 'input.xlsx', iterates through its worksheets, extracts each sheet's PaperSize, FitToPagesWide, and FitToPagesTall via the PageSetup object, writes these details into a new workbook 'PaperSizeReport.xlsx', and saves the report.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "PaperSizeReport.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(inputPath);

            // Create a new workbook for the report
            Workbook reportWorkbook = new Workbook();
            reportWorkbook.Worksheets.Clear(); // Remove the default sheet
            Worksheet reportSheet = reportWorkbook.Worksheets.Add("PaperSizeReport");

            // Write header row
            reportSheet.Cells[0, 0].PutValue("Worksheet");
            reportSheet.Cells[0, 1].PutValue("PaperSize");
            reportSheet.Cells[0, 2].PutValue("FitToPagesWide");
            reportSheet.Cells[0, 3].PutValue("FitToPagesTall");

            int reportRow = 1;

            // Iterate through each worksheet in the source workbook
            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                PageSetup pageSetup = ws.PageSetup;

                // Populate the report row with required values
                reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
                reportSheet.Cells[reportRow, 1].PutValue(pageSetup.PaperSize.ToString());
                reportSheet.Cells[reportRow, 2].PutValue(pageSetup.FitToPagesWide);
                reportSheet.Cells[reportRow, 3].PutValue(pageSetup.FitToPagesTall);

                reportRow++;
            }

            // Adjust column widths for better readability
            reportSheet.AutoFitColumns();

            // Save the report workbook
            reportWorkbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
