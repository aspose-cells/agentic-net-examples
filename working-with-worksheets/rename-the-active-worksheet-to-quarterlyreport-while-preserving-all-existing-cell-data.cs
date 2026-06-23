using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (contains a default worksheet)
        Workbook workbook = new Workbook();

        // Get the currently active worksheet via ActiveSheetIndex
        Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

        // Example data to show that cell contents are preserved after renaming
        activeSheet.Cells["A1"].PutValue("Sample Data");

        // Rename the active worksheet to "QuarterlyReport"
        activeSheet.Name = "QuarterlyReport";

        // Save the workbook; all existing cell data remains intact
        workbook.Save("QuarterlyReport.xlsx", SaveFormat.Xlsx);
    }
}