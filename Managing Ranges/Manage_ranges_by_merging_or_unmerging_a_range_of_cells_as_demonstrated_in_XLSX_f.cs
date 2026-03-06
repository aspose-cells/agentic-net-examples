using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Define a range covering cells A1:B2
        Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1", "B2");

        // Fill the range with sample data before merging
        range[0, 0].PutValue("Before Merge");
        range[0, 1].PutValue("A1");
        range[1, 0].PutValue("B1");
        range[1, 1].PutValue("B2");

        // Merge the cells in the range
        range.Merge();

        // Set a value in the merged cell
        range[0, 0].PutValue("Merged A1:B2");

        // Unmerge the previously merged cells
        range.UnMerge();

        // Populate individual cells after unmerging to demonstrate the effect
        worksheet.Cells["A1"].PutValue("A1");
        worksheet.Cells["A2"].PutValue("A2");
        worksheet.Cells["B1"].PutValue("B1");
        worksheet.Cells["B2"].PutValue("B2");

        // Save the workbook to an XLSX file
        workbook.Save("MergeUnmergeDemo.xlsx");
    }
}