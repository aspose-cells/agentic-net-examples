// Title: Configure every worksheet to print cell comments in place with Aspose.Cells for .NET (C#)
// AI Prompts: Iterate through workbook.Worksheets and set each sheet.PageSetup.PrintComments to PrintCommentsType.PrintInPlace, then save the file. | Apply in‑place comment printing to all worksheets in a C# Aspose.Cells project by updating the PrintComments property. | Enable printing of Excel comments exactly where they appear for every sheet and persist the changes using workbook.Save.
// Common Searches: Aspose.Cells how to print Excel comments in their original location for all sheets | C# set PrintCommentsType.PrintInPlace for every worksheet in a workbook | example code to enable in‑place comment printing with Aspose.Cells | change page setup to print comments in place using Aspose.Cells .NET | save workbook after modifying PrintComments property in Aspose.Cells
// Tags: Aspose.Cells PrintCommentsType.PrintInPlace | C# set worksheet PageSetup PrintComments | print Excel comments in place Aspose.Cells | modify workbook comment printing behavior | save workbook after changing PrintComments

using Aspose.Cells;

// Loads input.xlsx, sets each worksheet's PageSetup.PrintComments to PrintCommentsType.PrintInPlace so comments are printed where they appear, and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Load an existing workbook (uses the provided load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Set the PrintComments property to print comments in place for every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // PrintCommentsType.PrintInPlace prints comments exactly where they appear in the sheet
            sheet.PageSetup.PrintComments = PrintCommentsType.PrintInPlace;
        }

        // Save the modified workbook (uses the provided save rule)
        workbook.Save("output.xlsx");
    }
}
