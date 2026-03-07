using Aspose.Cells;
using System;

class RemoveSpecificPageBreak
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample horizontal page breaks
        worksheet.HorizontalPageBreaks.Add(5);   // break after row 5
        worksheet.HorizontalPageBreaks.Add(10);  // break after row 10
        worksheet.HorizontalPageBreaks.Add(15);  // break after row 15

        // Row index of the page break to remove (zero‑based)
        int targetRow = 10;

        // Locate the index of the page break with the specified row
        int removeIndex = -1;
        for (int i = 0; i < worksheet.HorizontalPageBreaks.Count; i++)
        {
            if (worksheet.HorizontalPageBreaks[i].Row == targetRow)
            {
                removeIndex = i;
                break;
            }
        }

        // Remove the page break if it exists
        if (removeIndex >= 0)
        {
            worksheet.HorizontalPageBreaks.RemoveAt(removeIndex);
        }

        // Save the workbook
        workbook.Save("RemovedPageBreak.xlsx");
    }
}