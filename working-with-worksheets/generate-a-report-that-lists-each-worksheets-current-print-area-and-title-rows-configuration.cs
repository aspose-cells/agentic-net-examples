// Title: Generate a PrintAreaReport worksheet that lists each sheet’s print area and title rows using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells in C# to add a new worksheet named 'PrintAreaReport' that records the PageSetup.PrintArea and PageSetup.PrintTitleRows of every other worksheet in an existing workbook and then save the file. | Write C# code that iterates through all worksheets, extracts their print area and title rows settings, writes the values into a summary sheet, and applies AutoFitColumns with Aspose.Cells. | Create a .NET program that loads an Excel file, generates a report sheet showing each worksheet’s current print area range and title rows configuration, and outputs the modified workbook.
// Common Searches: Aspose.Cells how to list print area for each worksheet in a workbook | C# retrieve print title rows of all sheets using Aspose.Cells | generate summary sheet with page setup information in .NET Aspose.Cells | auto fit columns after writing data with Aspose.Cells C# | save workbook after adding report worksheet Aspose.Cells
// Tags: Aspose.Cells C# extract worksheet print area | Aspose.Cells C# retrieve print title rows | Aspose.Cells generate print settings report sheet | Aspose.Cells auto-fit columns after data insertion | Aspose.Cells save workbook with added report worksheet

using Aspose.Cells;
using System;

// The example loads an existing workbook, adds a new worksheet called "PrintAreaReport", writes headers, then loops through all other worksheets to read each sheet's PageSetup.PrintArea and PageSetup.PrintTitleRows. It records the values (or "Not Set" when absent) in the report sheet, auto‑fits the columns for readability, and saves the workbook as a new file containing the generated report.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Add a new worksheet that will contain the report
        int reportSheetIndex = workbook.Worksheets.Add();
        Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
        reportSheet.Name = "PrintAreaReport";

        // Write header titles
        reportSheet.Cells["A1"].PutValue("Worksheet Name");
        reportSheet.Cells["B1"].PutValue("Print Area");
        reportSheet.Cells["C1"].PutValue("Title Rows");

        int currentRow = 1; // Zero‑based index; row 2 in Excel (first data row)

        // Iterate through all worksheets except the report sheet itself
        foreach (Worksheet ws in workbook.Worksheets)
        {
            if (ws.Name == reportSheet.Name)
                continue;

            // Retrieve the current print area (e.g., "$A$1:$D$20") – empty if not set
            string printArea = ws.PageSetup.PrintArea;

            // Retrieve the title rows configuration (e.g., "$1:$2") – empty if not set
            string titleRows = ws.PageSetup.PrintTitleRows;

            // Populate the report row
            reportSheet.Cells[currentRow, 0].PutValue(ws.Name);
            reportSheet.Cells[currentRow, 1].PutValue(string.IsNullOrEmpty(printArea) ? "Not Set" : printArea);
            reportSheet.Cells[currentRow, 2].PutValue(string.IsNullOrEmpty(titleRows) ? "Not Set" : titleRows);

            currentRow++;
        }

        // Adjust column widths for readability
        reportSheet.AutoFitColumns();

        // Save the workbook with the newly added report sheet
        workbook.Save("output_with_report.xlsx");
    }
}
