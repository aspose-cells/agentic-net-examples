using System;
using Aspose.Cells;

class MergeWorkbooks
{
    static void Main()
    {
        // Create the destination workbook
        Workbook mergedWorkbook = new Workbook();
        mergedWorkbook.Worksheets.Clear(); // remove the default sheet

        // Load source workbooks and copy their worksheets into the merged workbook
        string[] sourceFiles = { "Source1.xlsx", "Source2.xlsx" };
        foreach (string file in sourceFiles)
        {
            Workbook src = new Workbook(file);
            foreach (Worksheet ws in src.Worksheets)
            {
                int newIndex = mergedWorkbook.Worksheets.Add();
                Worksheet target = mergedWorkbook.Worksheets[newIndex];
                ws.Copy(target);
            }
        }

        // Set the author of the merged workbook
        mergedWorkbook.Settings.Author = Environment.UserName;

        // Save the merged workbook
        mergedWorkbook.Save("MergedOutput.xlsx");
    }
}