// Title: C# – Aspose.Cells: Generate a Report of Each Worksheet’s Print Area and Title Rows
// Description: Loads an existing workbook, adds a "PrintAreaReport" sheet, and iterates through all worksheets (except the report sheet) to capture PageSetup.PrintArea and PageSetup.PrintTitleRows. The code writes the worksheet name, its print area, and title rows (or "Not Set") into the new sheet and saves the file as PrintAreaReport.xlsx.
// Keywords: Aspose.Cells | C# | print area | print title rows | worksheet report | PageSetup | export print settings | Excel automation | list worksheets | summary sheet
// Common Searches: Aspose.Cells get print area for each sheet | C# list worksheet print title rows Aspose.Cells | Create Excel report of print settings with .NET | How to export print area to a new worksheet using Aspose | Generate summary of worksheet print configuration
// Developer Intent: Create a summary worksheet that lists every sheet’s defined print area and print title rows.
// Use Cases: Audit workbook print configurations before batch printing to verify correct ranges. | Provide end‑users with a quick reference of print settings for documentation or troubleshooting. | Validate that all generated worksheets contain required print areas and title rows.
// AI Prompts: Write C# code with Aspose.Cells that adds a summary sheet showing each worksheet’s PrintArea and PrintTitleRows. | Extend the example to also capture PrintTitleColumns and include them as a fourth column in the report. | Add error handling for missing input files and give an option to export the summary as PDF instead of XLSX.

using System;
using Aspose.Cells;

// Loads an existing workbook, adds a "PrintAreaReport" sheet, and iterates through all worksheets (except the report sheet) to capture PageSetup.PrintArea and PageSetup.PrintTitleRows. The code writes the worksheet name, its print area, and title rows (or "Not Set") into the new sheet and saves the file as PrintAreaReport.xlsx.
class PrintAreaReport
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your actual file)
        Workbook workbook = new Workbook("input.xlsx");

        // Add a new worksheet that will contain the report
        int reportSheetIndex = workbook.Worksheets.Add();
        Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
        reportSheet.Name = "PrintAreaReport";

        // Write column headers
        reportSheet.Cells[0, 0].PutValue("Worksheet");
        reportSheet.Cells[0, 1].PutValue("Print Area");
        reportSheet.Cells[0, 2].PutValue("Print Title Rows");

        // Populate the report with each worksheet's print settings
        int currentRow = 1;
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Skip the report sheet itself
            if (ws.Index == reportSheetIndex) continue;

            string printArea = ws.PageSetup.PrintArea;
            string titleRows = ws.PageSetup.PrintTitleRows;

            reportSheet.Cells[currentRow, 0].PutValue(ws.Name);
            reportSheet.Cells[currentRow, 1].PutValue(string.IsNullOrEmpty(printArea) ? "Not Set" : printArea);
            reportSheet.Cells[currentRow, 2].PutValue(string.IsNullOrEmpty(titleRows) ? "Not Set" : titleRows);
            currentRow++;
        }

        // Save the workbook containing the report
        workbook.Save("PrintAreaReport.xlsx");
    }
}
