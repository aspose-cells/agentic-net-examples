using System;
using Aspose.Cells;

class MergeWorkbooksWithTitle
{
    static void Main()
    {
        // Paths of the source workbooks to be merged
        string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx" };

        // Create an empty destination workbook
        Workbook mergedWorkbook = new Workbook();

        // Iterate through each source file, load it and combine into the destination workbook
        foreach (string filePath in sourceFiles)
        {
            Workbook source = new Workbook(filePath);
            mergedWorkbook.Combine(source);
        }

        // Assign a descriptive title that reflects the combined source files
        mergedWorkbook.BuiltInDocumentProperties.Title = "Combined Workbook: Source1.xlsx + Source2.xlsx";

        // Save the merged workbook to disk
        mergedWorkbook.Save("MergedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}