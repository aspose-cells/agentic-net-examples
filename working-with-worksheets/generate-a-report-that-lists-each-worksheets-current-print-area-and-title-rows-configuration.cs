// Title: Generate a Print Area and Title Rows Report Sheet for an Excel Workbook using Aspose.Cells for .NET
// Common Searches: how to list print area of each worksheet with Aspose.Cells in C# | C# Aspose.Cells code to retrieve worksheet title rows settings | create a summary sheet that shows Excel print settings using Aspose.Cells | Aspose.Cells iterate worksheets and export page setup information to new sheet | skip the report worksheet while generating print area report with Aspose.Cells
// Tags: Aspose.Cells get worksheet print area | Aspose.Cells read print title rows | Aspose.Cells generate print settings summary sheet | Aspose.Cells exclude specific worksheet during iteration | Aspose.Cells auto-fit columns after data insertion | Aspose.Cells save workbook with additional report sheet

using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaReport
{
    // The example loads an existing workbook, adds a new worksheet named "PrintAreaReport", writes header cells, then loops through all other worksheets to capture each sheet's PageSetup.PrintArea and PageSetup.PrintTitleRows (using "Not Set" when undefined). It records the worksheet name, print area, and title rows in the report sheet, auto‑fits the columns for readability, and saves the workbook as a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the source workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Create (or reuse) a worksheet to hold the report
            Worksheet reportSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            reportSheet.Name = "PrintAreaReport";

            // Write header row
            reportSheet.Cells[0, 0].PutValue("Worksheet Name");
            reportSheet.Cells[0, 1].PutValue("Print Area");
            reportSheet.Cells[0, 2].PutValue("Title Rows");

            int reportRow = 1; // start writing data from the second row

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Skip the report sheet itself to avoid self‑reference
                if (sheet.Name == reportSheet.Name)
                    continue;

                // Retrieve the current print area (if any)
                string printArea = sheet.PageSetup.PrintArea; // e.g., "$A$1:$D$20"
                if (string.IsNullOrEmpty(printArea))
                    printArea = "Not Set";

                // Retrieve the title rows configuration (if any)
                string titleRows = sheet.PageSetup.PrintTitleRows; // e.g., "$1:$2"
                if (string.IsNullOrEmpty(titleRows))
                    titleRows = "Not Set";

                // Write the information into the report sheet
                reportSheet.Cells[reportRow, 0].PutValue(sheet.Name);
                reportSheet.Cells[reportRow, 1].PutValue(printArea);
                reportSheet.Cells[reportRow, 2].PutValue(titleRows);

                reportRow++;
            }

            // Auto‑fit columns for better readability
            reportSheet.AutoFitColumns();

            // Save the workbook with the new report sheet
            workbook.Save("output.xlsx");
        }
    }
}
