using System;
using Aspose.Cells;

namespace AsposeCellsPrintCommentsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Audit note: verify this value.";

            // Configure the worksheet to print comments as separate notes at the end of the sheet
            // This uses the PrintComments property of PageSetup with the PrintSheetEnd option
            sheet.PageSetup.PrintComments = PrintCommentsType.PrintSheetEnd;

            // Save the workbook (save rule)
            workbook.Save("AuditCommentsPrint.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved with comments set to print at the end of the sheet.");
        }
    }
}