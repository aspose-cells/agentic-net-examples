// Title: Aspose.Cells for .NET – Print worksheet comments as separate notes at the sheet end (C#)
// Description: Demonstrates how to create a workbook, add a comment, and configure PageSetup.PrintComments to PrintCommentsType.PrintSheetEnd so that all comments are printed as independent notes at the end of the worksheet—ideal for audit‑ready reports. The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells C# print comments | PrintCommentsType.PrintSheetEnd | worksheet comments as notes | Excel audit comments | page setup print comments | export comments to PDF Aspose.Cells | .NET Excel comment printing | separate comment notes Excel
// Common Searches: Aspose.Cells print comments at end of sheet C# | How to output Excel comments as separate notes using Aspose.Cells | PrintCommentsType.PrintSheetEnd example | Set page setup to print comments after data in .NET | Generate audit report with comments printed at sheet end
// Developer Intent: Configure a worksheet so that its cell comments are printed as independent notes at the end of the sheet.
// Use Cases: Create audit‑ready spreadsheets where all reviewer notes are collected in a single section. | Produce printable reports that keep data cells clean while still providing full commentary. | Generate documentation that consolidates user feedback after the main content for easier review.
// AI Prompts: Show a C# snippet that sets PrintComments = PrintCommentsType.PrintSheetEnd for every worksheet and saves the workbook as PDF. | Explain the differences between PrintCommentsType.PrintSheetEnd and PrintCommentsType.PrintInPlace when exporting to various formats. | Provide code to iterate through a workbook, apply the sheet‑end comment printing setting, and then export each sheet to a separate file.

using System;
using Aspose.Cells;

namespace AsposeCellsCommentPrintDemo
{
    // Demonstrates how to create a workbook, add a comment, and configure PageSetup.PrintComments to PrintCommentsType.PrintSheetEnd so that all comments are printed as independent notes at the end of the worksheet—ideal for audit‑ready reports. The workbook is saved as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");

            // Add a comment to a cell
            int commentIdx = sheet.Comments.Add("A2");
            Comment comment = sheet.Comments[commentIdx];
            comment.Note = "This comment will be printed at the end of the sheet for audit.";

            // Configure the page setup to print comments as separate notes (at sheet end)
            sheet.PageSetup.PrintComments = PrintCommentsType.PrintSheetEnd;

            // Save the workbook (lifecycle rule: save)
            workbook.Save("CommentsPrintedAtSheetEnd.xlsx", SaveFormat.Xlsx);
        }
    }
}
