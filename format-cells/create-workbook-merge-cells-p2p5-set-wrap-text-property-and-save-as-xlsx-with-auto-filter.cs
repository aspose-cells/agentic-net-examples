// Title: Merge cells P2:P5, enable text wrap, add AutoFilter, and save as XLSX using Aspose.Cells for .NET (C#)
// Common Searches: How to merge cells P2:P5 and enable text wrap with Aspose.Cells in C# | Aspose.Cells .NET code to add an AutoFilter to the first row after merging cells | Saving a workbook that contains merged and wrapped cells as an XLSX file using Aspose.Cells | C# example for merging column P rows 2 through 5 with Aspose.Cells | Applying wrap text to a merged cell and setting AutoFilter with Aspose.Cells
// Tags: merge cells range Aspose.Cells | wrap text style merged cell Aspose.Cells | apply autofilter to header row Aspose.Cells | save workbook as xlsx Aspose.Cells | create workbook with merged cells Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new workbook, merges the range P2:P5, sets the merged cell to wrap text, adds header values, enables an AutoFilter on the header row, and saves the result as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells P2:P5 (column P = index 15, rows 2-5 = indices 1-4).
            sheet.Cells.Merge(1, 15, 4, 1);

            // Set wrap text for the merged cell.
            Style mergedStyle = sheet.Cells[1, 15].GetStyle();
            mergedStyle.IsTextWrapped = true;
            sheet.Cells[1, 15].SetStyle(mergedStyle);

            // Add sample headers.
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["C1"].PutValue("Header3");

            // Enable AutoFilter on the header row.
            sheet.AutoFilter.Range = "A1:C1";

            // Save the workbook as XLSX.
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
