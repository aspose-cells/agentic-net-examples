// Title: Insert a Merge Timestamp Worksheet After Combining Excel Workbooks with Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells to add a new worksheet named 'MergeTimestamp' after calling Workbook.Combine, then write DateTime.Now into cell A1. | Apply Excel's built‑in date‑time number format (style number 22) to the timestamp cell and save the merged workbook as a separate file.
// Common Searches: how to add a timestamp sheet after merging two Excel files using Aspose.Cells C# | record merge date and time in a new worksheet with Aspose.Cells .NET | Aspose.Cells combine workbooks and create a log worksheet with current datetime | C# Aspose.Cells add worksheet and set built‑in date‑time format after workbook combine | save merged workbook with additional timestamp worksheet using Aspose.Cells
// Tags: merge workbooks add timestamp worksheet Aspose.Cells | write current datetime to cell A1 Aspose.Cells | apply built‑in date‑time style number 22 Aspose.Cells | save combined workbook with extra sheet C# | Workbook.Combine post‑merge logging Aspose.Cells

using System;
using Aspose.Cells;

// The example loads two Excel files, merges the source workbook into the destination using Workbook.Combine, creates a new worksheet called 'MergeTimestamp', writes the current DateTime into cell A1, formats the cell with Excel's built‑in date‑time style, and saves the result as a new file containing the timestamp sheet.
class MergeWithTimestamp
{
    static void Main()
    {
        // Load the workbooks to be merged
        Workbook sourceWorkbook = new Workbook("Source.xlsx");
        Workbook destinationWorkbook = new Workbook("Destination.xlsx");

        // Merge the source workbook into the destination workbook
        destinationWorkbook.Combine(sourceWorkbook);

        // Add a new worksheet that will hold the merge timestamp
        Worksheet timestampSheet = destinationWorkbook.Worksheets.Add("MergeTimestamp");

        // Write the exact date and time of the merge operation into cell A1
        timestampSheet.Cells["A1"].PutValue(DateTime.Now);

        // Apply a standard date‑time number format to the cell (optional)
        Style dateTimeStyle = destinationWorkbook.CreateStyle();
        dateTimeStyle.Number = 22; // Built‑in Excel date‑time format
        timestampSheet.Cells["A1"].SetStyle(dateTimeStyle);

        // Save the merged workbook with the timestamp worksheet
        destinationWorkbook.Save("MergedWithTimestamp.xlsx");
    }
}
