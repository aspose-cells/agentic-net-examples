using System;
using Aspose.Cells;

class MergeWorkbooks
{
    static void Main()
    {
        // Paths of the workbooks to be merged
        string[] sourceFiles = new string[]
        {
            "File1.xlsx",
            "File2.xlsx",
            "File3.xlsx"
        };

        // Create an empty destination workbook
        Workbook combinedWorkbook = new Workbook();

        // Iterate over each source workbook and combine it into the destination
        foreach (string filePath in sourceFiles)
        {
            // Load the source workbook (preserves all data, formatting, and formulas)
            Workbook sourceWorkbook = new Workbook(filePath);

            // Merge the source workbook into the destination workbook
            combinedWorkbook.Combine(sourceWorkbook);
        }

        // Save the merged workbook to disk (all original formatting and formulas are retained)
        combinedWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}