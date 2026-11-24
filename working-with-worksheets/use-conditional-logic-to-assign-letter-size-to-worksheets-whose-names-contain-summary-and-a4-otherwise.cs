// Title: Set Letter paper size for worksheets whose names contain "Summary" and A4 for all other sheets with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set paper size based on worksheet name | How to apply Letter size to summary worksheets using Aspose.Cells | Conditional page setup for multiple sheets in a .NET workbook | Set default A4 paper size for all worksheets except specific ones Aspose.Cells | Iterate through worksheets and change PageSetup.PaperSize in C#
// Tags: conditional worksheet paper size Aspose.Cells | set page setup by sheet name C# | letter paper size for summary sheets | default a4 page setup Aspose.Cells | iterate worksheets set paper size .NET

using Aspose.Cells;

// The code loads a workbook, iterates over each worksheet, assigns PaperLetter when the sheet name contains "Summary" and PaperA4 otherwise, and then saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // If the worksheet name contains "Summary", set paper size to Letter
            if (sheet.Name.Contains("Summary"))
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
            }
            // Otherwise, set paper size to A4
            else
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
