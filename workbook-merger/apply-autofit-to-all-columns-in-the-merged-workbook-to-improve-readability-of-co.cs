using System;
using Aspose.Cells;

class AutoFitMergedWorkbook
{
    static void Main()
    {
        // Create an empty workbook that will hold the combined data
        Workbook mergedWorkbook = new Workbook();

        // Define the source workbook files to be merged
        string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx" };

        // Load each source workbook and combine it into the merged workbook
        foreach (string filePath in sourceFiles)
        {
            Workbook source = new Workbook(filePath);   // Load source workbook
            mergedWorkbook.Combine(source);            // Merge into the destination workbook
        }

        // AutoFit all columns in every worksheet of the merged workbook
        foreach (Worksheet sheet in mergedWorkbook.Worksheets)
        {
            sheet.AutoFitColumns();                    // Adjust column widths for readability
        }

        // Save the final merged workbook with auto‑fitted columns
        mergedWorkbook.Save("MergedAutoFit.xlsx", SaveFormat.Xlsx);
    }
}