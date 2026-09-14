// Title: How to set PrintComments to PrintInPlace for every worksheet in an Aspose.Cells workbook using C#
// AI Prompts: Write C# code that loads a workbook, iterates through all worksheets, and assigns PageSetup.PrintComments = PrintCommentsType.PrintInPlace before saving the file. | Show a complete example of configuring Aspose.Cells to print cell comments at their original positions for all sheets and then persisting the changes.
// Common Searches: Aspose.Cells C# set page setup to print comments in place for all sheets | Print cell comments exactly where they appear using Aspose.Cells .NET | How to apply PrintCommentsType.PrintInPlace to every worksheet in a workbook | C# code example for changing PrintComments property in Aspose.Cells | Saving workbook after modifying PrintComments setting with Aspose.Cells
// Tags: Aspose.Cells PrintCommentsType.PrintInPlace configuration | C# loop worksheets set PageSetup.PrintComments | Aspose.Cells page setup comment printing | Workbook.Save after page setup modification | In‑place comment printing Aspose.Cells .NET

using System;
using Aspose.Cells;

// The sample loads an existing workbook, loops through each worksheet, sets PageSetup.PrintComments to PrintCommentsType.PrintInPlace so comments are printed at their original locations, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets and set the PrintComments option
        // to print comments in place on the printed page.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // PrintComments is a property of PageSetup.
            // PrintCommentsType.PrintInPlace prints the comment exactly where it appears in the sheet.
            sheet.PageSetup.PrintComments = PrintCommentsType.PrintInPlace;
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
