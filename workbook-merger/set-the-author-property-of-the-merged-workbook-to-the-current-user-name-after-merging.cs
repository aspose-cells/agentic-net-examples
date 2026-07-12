using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create the first workbook (could be loaded from a file instead)
        Workbook firstWorkbook = new Workbook();
        firstWorkbook.Worksheets[0].Cells["A1"].PutValue("Data from first workbook");

        // Create the second workbook (could be loaded from a file instead)
        Workbook secondWorkbook = new Workbook();
        secondWorkbook.Worksheets[0].Cells["A1"].PutValue("Data from second workbook");

        // Merge the second workbook into the first one
        firstWorkbook.Combine(secondWorkbook);

        // Set the author of the merged workbook to the current user name
        firstWorkbook.Settings.Author = Environment.UserName;

        // Save the merged workbook
        firstWorkbook.Save("MergedWorkbook.xlsx");
    }
}