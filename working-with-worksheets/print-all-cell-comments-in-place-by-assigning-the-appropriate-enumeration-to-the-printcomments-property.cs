// Title: Print cell comments in place with Aspose.Cells for .NET (PrintComments = PrintInPlace)
// Description: Shows how to add comments to cells, configure Worksheet.PageSetup.PrintComments to PrintCommentsType.PrintInPlace, and save the workbook so comments retain their on‑sheet positions when printed or exported.
// Keywords: Aspose.Cells | PrintComments | PrintInPlace | C# cell comments | worksheet page setup | export comments | .NET spreadsheet printing | cell notes printing
// Common Searches: Aspose.Cells print comments in place C# | How to keep cell comments when printing with Aspose.Cells | Set PrintComments property Aspose.Cells | Print cell notes exactly as shown Aspose.Cells | PageSetup.PrintComments example .NET
// Developer Intent: Configure the worksheet so that all cell comments are printed at their original locations on the sheet.
// Use Cases: Create printable reports where reviewer notes stay anchored to their cells. | Generate audit worksheets that must display comments on hard‑copy output. | Export a workbook to PDF while preserving the visual layout of cell comments.
// AI Prompts: Provide C# code that adds comments to multiple cells, sets PrintComments to PrintInPlace for every worksheet, and saves the workbook as PDF. | Explain the difference between PrintInPlace and PrintNoComments in Aspose.Cells and recommend scenarios for each. | Show how to apply PageSetup.PrintComments = PrintInPlace to all worksheets in an existing workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to add comments to cells, configure Worksheet.PageSetup.PrintComments to PrintCommentsType.PrintInPlace, and save the workbook so comments retain their on‑sheet positions when printed or exported.
    public class PrintAllCommentsInPlaceDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample comments to demonstrate the setting
            int idx1 = sheet.Comments.Add("A1");
            sheet.Comments[idx1].Note = "Comment for A1";

            int idx2 = sheet.Comments.Add("B2");
            sheet.Comments[idx2].Note = "Comment for B2";

            // Set the page setup to print comments exactly as they appear on the sheet
            sheet.PageSetup.PrintComments = PrintCommentsType.PrintInPlace;

            // Save the workbook (the comments will be printed in place when the sheet is printed)
            string outputPath = "PrintCommentsInPlace.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}' with PrintComments set to PrintInPlace.");
        }
    }
}
