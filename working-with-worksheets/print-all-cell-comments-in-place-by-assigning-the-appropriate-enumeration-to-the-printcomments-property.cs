// Title: Print Cell Comments In‑Place with Aspose.Cells for .NET (C#)
// Description: Shows how to add comments to cells, set Worksheet.PageSetup.PrintComments to PrintCommentsType.PrintInPlace, and save the workbook so the comments are printed directly on the sheet.
// Keywords: Aspose.Cells | PrintComments | PrintInPlace | C# | .NET | worksheet comments | print comments in place | Excel export | PageSetup | SaveFormat.Xlsx
// Common Searches: Aspose.Cells print comments in place C# | Set PrintComments property Aspose.Cells .NET | How to print cell comments on worksheet using Aspose.Cells | PrintCommentsType.PrintInPlace example | Export Excel with visible comments Aspose.Cells
// Developer Intent: Configure a worksheet to print its cell comments directly on the printed page by using the PrintComments property.
// Use Cases: Create printable Excel reports where reviewer notes stay attached to their cells. | Distribute workbooks that retain visible comments on paper copies. | Automate generation of documentation sheets with in‑place comments for quality checks.
// AI Prompts: Provide C# code to set PrintComments = PrintInPlace for all worksheets in a workbook using Aspose.Cells. | Show how to export a workbook with in‑place comments to PDF with Aspose.Cells. | Explain the differences between PrintInPlace, PrintNoComments, and PrintAllComments options in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintCommentsDemo
{
    // Shows how to add comments to cells, set Worksheet.PageSetup.PrintComments to PrintCommentsType.PrintInPlace, and save the workbook so the comments are printed directly on the sheet.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample comments to demonstrate the print setting
            int idx1 = sheet.Comments.Add("A1");
            sheet.Comments[idx1].Note = "Comment for A1";

            int idx2 = sheet.Comments.Add("B2");
            sheet.Comments[idx2].Note = "Comment for B2";

            // Set the PrintComments property to print comments in place
            sheet.PageSetup.PrintComments = PrintCommentsType.PrintInPlace;

            // Save the workbook (the comments will be printed in place when printed)
            workbook.Save("PrintCommentsInPlace.xlsx", SaveFormat.Xlsx);
        }
    }
}
