// Author: Aspose.Cells .NET example – export a large workbook with many comments efficiently
using System;
using Aspose.Cells;

class ExportCommentsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a large number of comments to simulate an extensive comment scenario
        for (int row = 0; row < 5000; row++)
        {
            // Add comment to column A of each row
            int commentIndex = sheet.Comments.Add(row, 0);
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = $"Comment for cell A{row + 1}";
        }

        // Configure OoxmlSaveOptions to improve performance for large files
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            // Disable exporting cell names – reduces file size and speeds up saving
            ExportCellName = false
        };

        // Save the workbook as XLSX using the configured options
        workbook.Save("LargeCommentsWorkbook.xlsx", saveOptions);
    }
}