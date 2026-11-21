// Title: How to set black‑and‑white printing mode for all worksheets in an Aspose.Cells workbook using C#
// Common Searches: Aspose.Cells C# set black and white printing for every worksheet in a workbook | Enable grayscale printing for all sheets with Aspose.Cells API | Reduce ink usage by configuring PageSetup.BlackAndWhite in C# Excel export | Apply black‑and‑white page setup to all worksheets programmatically | Iterate worksheets to set BlackAndWhite property in Aspose.Cells
// Tags: Aspose.Cells PageSetup.BlackAndWhite C# | grayscale printing for entire workbook | bulk worksheet print configuration Aspose.Cells | ink saving Excel export Aspose.Cells | programmatic black‑and‑white mode workbook

using Aspose.Cells;

// The code creates a Workbook, iterates through each Worksheet, sets the PageSetup.BlackAndWhite property to true to enforce black‑and‑white (grayscale) printing for the whole workbook, and then saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (replace with load if needed)
        Workbook workbook = new Workbook();

        // Enable black‑and‑white printing for every worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.PageSetup.BlackAndWhite = true;
        }

        // Save the workbook to a file
        workbook.Save("BlackAndWhiteWorkbook.xlsx");
    }
}
