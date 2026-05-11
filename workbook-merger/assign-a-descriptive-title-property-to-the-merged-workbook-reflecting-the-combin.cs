using System;
using Aspose.Cells;

class MergeWorkbooksWithTitle
{
    static void Main()
    {
        // Paths of source workbooks to be merged
        string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx", "Source3.xlsx" };

        // Create an empty destination workbook
        Workbook mergedWorkbook = new Workbook();

        // Merge each source workbook into the destination workbook
        foreach (string filePath in sourceFiles)
        {
            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(filePath);

            // Combine the source workbook with the merged workbook
            mergedWorkbook.Combine(sourceWorkbook);
        }

        // Set a descriptive title reflecting the combined source files
        mergedWorkbook.BuiltInDocumentProperties.Title =
            "Combined Workbook: " + string.Join(", ", sourceFiles);

        // Save the merged workbook to a file
        mergedWorkbook.Save("MergedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}