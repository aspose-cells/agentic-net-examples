using Aspose.Cells;
using System;

class MergeUnmergeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells A1:C3 (rows 0-2, columns 0-2)
        cells.Merge(0, 0, 3, 3);
        // Put a value in the merged cell (upper‑left corner)
        cells[0, 0].PutValue("Merged A1:C3");

        // Merge another range D5:E6 (rows 4-5, columns 3-4)
        cells.Merge(4, 3, 2, 2);
        cells[4, 3].PutValue("Merged D5:E6");

        // Unmerge the first range A1:C3
        cells.UnMerge(0, 0, 3, 3);
        // After unmerging, populate individual cells to demonstrate the result
        cells["A1"].PutValue("A1");
        cells["B2"].PutValue("B2");
        cells["C3"].PutValue("C3");

        // Save the workbook to an XLSX file
        workbook.Save("MergeUnmergeDemo.xlsx");
    }
}