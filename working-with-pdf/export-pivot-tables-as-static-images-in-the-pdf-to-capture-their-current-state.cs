// Title: Export Excel Pivot Tables as Static Images into a PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, refreshes all pivot tables, copies each worksheet containing a pivot table into a new workbook, and saves the new workbook as a PDF with Aspose.Cells. | Adjust the Aspose.Cells example to skip pivot tables that cannot be refreshed, log their worksheet and table names, and continue processing the remaining tables. | Create a script that renames each copied worksheet with the original sheet and pivot table names, then exports the collection of worksheets to a single PDF file.
// Common Searches: Aspose.Cells .NET export refreshed pivot tables to PDF as images | C# code to convert Excel pivot tables into static PDF pages using Aspose.Cells | How to copy worksheets with pivot tables and save them as PDF with Aspose.Cells | Save Excel pivot table view as image in PDF using Aspose.Cells for .NET | Skip problematic pivot tables while exporting to PDF with Aspose.Cells
// Tags: Aspose.Cells pivot table PDF conversion | refresh pivot cache before PDF export | duplicate worksheet with pivot table Aspose.Cells | static image rendering of pivot tables in PDF | robust pivot table export error handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an Excel workbook, refreshes each pivot table, copies the worksheets that contain them into a new workbook, and saves the new workbook as a PDF, producing static images of the pivot tables on separate pages.
class ExportPivotTablesAsImagesToPdf
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains the pivot tables
            Workbook srcWorkbook = new Workbook(inputPath);

            // Create a new workbook that will hold the pivot table copies
            Workbook pdfWorkbook = new Workbook();
            pdfWorkbook.Worksheets.Clear(); // remove the default sheet

            // Iterate through all worksheets and their pivot tables
            foreach (Worksheet ws in srcWorkbook.Worksheets)
            {
                foreach (PivotTable pt in ws.PivotTables)
                {
                    try
                    {
                        // Refresh pivot cache data and recalculate
                        pt.RefreshData();          // correct method to refresh pivot cache
                        pt.CalculateData();

                        // Add a copy of the worksheet containing this pivot table
                        int sheetIndex = pdfWorkbook.Worksheets.AddCopy(ws.Name);
                        Worksheet copiedSheet = pdfWorkbook.Worksheets[sheetIndex];
                        copiedSheet.Name = $"{ws.Name}_{pt.Name}";
                    }
                    catch (Exception exPivot)
                    {
                        Console.WriteLine($"Failed to process pivot table '{pt.Name}' in sheet '{ws.Name}': {exPivot.Message}");
                    }
                }
            }

            // Save the workbook as a PDF – each worksheet will appear as a page
            pdfWorkbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
