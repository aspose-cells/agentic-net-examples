using System;
using Aspose.Cells;

class PrintAreaReport
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Add a new worksheet that will contain the report
        int reportSheetIndex = workbook.Worksheets.Add();
        Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
        reportSheet.Name = "PrintAreaReport";

        // Write header row in the report sheet
        reportSheet.Cells[0, 0].PutValue("Worksheet");
        reportSheet.Cells[0, 1].PutValue("Print Area");
        reportSheet.Cells[0, 2].PutValue("Print Title Rows");

        int reportRow = 1; // start writing data from the second row

        // Iterate through all worksheets in the workbook
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Skip the report sheet itself to avoid recursion
            if (ws.Index == reportSheetIndex) continue;

            // Retrieve print area and title rows from the PageSetup object
            string printArea = ws.PageSetup.PrintArea;
            string titleRows = ws.PageSetup.PrintTitleRows;

            // Write the information into the report worksheet
            reportSheet.Cells[reportRow, 0].PutValue(ws.Name);
            reportSheet.Cells[reportRow, 1].PutValue(string.IsNullOrEmpty(printArea) ? "Not Set" : printArea);
            reportSheet.Cells[reportRow, 2].PutValue(string.IsNullOrEmpty(titleRows) ? "Not Set" : titleRows);

            // Also output the same information to the console
            Console.WriteLine($"Sheet: {ws.Name}, PrintArea: {printArea}, TitleRows: {titleRows}");

            reportRow++;
        }

        // Save the workbook with the newly added report sheet
        workbook.Save("output_with_report.xlsx");
    }
}