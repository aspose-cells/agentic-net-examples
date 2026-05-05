using System;
using Aspose.Cells;

namespace WorksheetManagementDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the worksheet collection
            WorksheetCollection sheets = workbook.Worksheets;

            // Add three worksheets with specific names
            Worksheet dataSheet = sheets.Add("Data");
            Worksheet summarySheet = sheets.Add("Summary");
            Worksheet archiveSheet = sheets.Add("Archive");

            // Rename the "Archive" worksheet to "OldData"
            archiveSheet.Name = "OldData";

            // Delete the "Summary" worksheet using its index
            sheets.RemoveAt(sheets["Summary"].Index);

            // Reorder worksheets: move "OldData" to the first position (index 0)
            Worksheet oldDataSheet = sheets["OldData"];
            oldDataSheet.MoveTo(0);

            // Hide the "OldData" worksheet
            oldDataSheet.SetVisible(false, false);

            // Save the workbook to disk
            workbook.Save("WorksheetManagementResult.xlsx");
        }
    }
}