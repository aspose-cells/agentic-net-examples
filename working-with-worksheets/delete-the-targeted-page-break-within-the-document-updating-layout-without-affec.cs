using System;
using Aspose.Cells;

class DeletePageBreakDemo
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample horizontal page breaks
        worksheet.HorizontalPageBreaks.Add(5);
        worksheet.HorizontalPageBreaks.Add(10);
        worksheet.HorizontalPageBreaks.Add(15);

        // Target horizontal page break at row 10
        int targetRow = 10;
        int hIndexToRemove = -1;
        for (int i = 0; i < worksheet.HorizontalPageBreaks.Count; i++)
        {
            if (worksheet.HorizontalPageBreaks[i].Row == targetRow)
            {
                hIndexToRemove = i;
                break;
            }
        }
        if (hIndexToRemove >= 0)
        {
            // Remove the identified horizontal page break
            worksheet.HorizontalPageBreaks.RemoveAt(hIndexToRemove);
        }

        // Add sample vertical page breaks
        worksheet.VerticalPageBreaks.Add(4);
        worksheet.VerticalPageBreaks.Add(7);
        worksheet.VerticalPageBreaks.Add(10);

        // Target vertical page break at column 7
        int targetColumn = 7;
        int vIndexToRemove = -1;
        for (int i = 0; i < worksheet.VerticalPageBreaks.Count; i++)
        {
            if (worksheet.VerticalPageBreaks[i].Column == targetColumn)
            {
                vIndexToRemove = i;
                break;
            }
        }
        if (vIndexToRemove >= 0)
        {
            // Remove the identified vertical page break
            worksheet.VerticalPageBreaks.RemoveAt(vIndexToRemove);
        }

        // Save the workbook with updated page breaks
        workbook.Save("output.xlsx");
    }
}