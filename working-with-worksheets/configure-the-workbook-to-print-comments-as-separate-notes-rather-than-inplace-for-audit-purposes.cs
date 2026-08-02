// Title: Aspose.Cells C# – Print worksheet comments as separate notes at the end of the sheet
// Description: Demonstrates how to add a comment to a cell, configure the worksheet's PageSetup to use PrintCommentsType.PrintSheetEnd, and save the workbook so that all comments are printed as independent notes after the data section—ideal for audit‑ready reports.
// Keywords: Aspose.Cells | C# | PrintComments | PrintSheetEnd | Excel comments as notes | page setup | audit report | separate comment printing
// Common Searches: Aspose.Cells print comments at sheet end | C# print Excel comments as separate notes | How to configure PrintCommentsType in Aspose.Cells | Export workbook with comments listed after data | Audit Excel file with comments printed separately
// Developer Intent: Configure a worksheet so that its comments are printed as distinct notes at the end of the sheet.
// Use Cases: Create audit‑ready Excel files where each comment appears in a consolidated notes section. | Generate printable reports that list all cell remarks after the main data for stakeholder review. | Meet compliance requirements by separating commentary from the data layout in exported workbooks.
// AI Prompts: Show C# code that sets PrintComments to PrintSheetEnd for a worksheet using Aspose.Cells. | Provide an example of printing comments as separate notes for multiple sheets in a workbook. | Explain how to verify the PrintComments setting before saving the Excel file.

using System;
using Aspose.Cells;

// Demonstrates how to add a comment to a cell, configure the worksheet's PageSetup to use PrintCommentsType.PrintSheetEnd, and save the workbook so that all comments are printed as independent notes after the data section—ideal for audit‑ready reports.
class PrintCommentsSeparateNotes
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a comment to cell A1 (this will be printed as a separate note)
        int commentIndex = sheet.Comments.Add("A1");
        Comment comment = sheet.Comments[commentIndex];
        comment.Note = "Audit note for cell A1";

        // Configure the page setup to print comments at the end of the sheet
        sheet.PageSetup.PrintComments = PrintCommentsType.PrintSheetEnd;

        // Save the workbook (the setting persists in the saved file)
        workbook.Save("AuditComments.xlsx", SaveFormat.Xlsx);
    }
}
