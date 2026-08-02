// Title: C# Aspose.Cells: Set worksheet page‑setup to A2 paper size and read dimensions
// Description: Step‑by‑step example showing how to assign PaperSizeType.PaperA2 to a worksheet, retrieve PaperWidth and PaperHeight in inches, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# set paper size | PaperA2 page setup | worksheet paper dimensions | PaperWidth Aspose.Cells | PaperHeight Aspose.Cells | A2 paper size .NET | save workbook Aspose.Cells | page setup API | retrieve paper size inches
// Common Searches: how to set worksheet paper size to A2 Aspose.Cells | read paper width and height after setting page setup | Aspose.Cells C# A2 paper dimensions | save workbook after changing page setup paper size | Aspose.Cells get paper size in inches
// Developer Intent: Set a worksheet’s page‑setup to A2 and confirm the resulting width and height values.
// Use Cases: Generate printable reports that must fit A2 sheets and verify dimensions before printing. | Validate spreadsheet templates against A2 specifications for distribution. | Automate creation of large‑format (A2) spreadsheets in a batch‑processing workflow.
// AI Prompts: Write C# code with Aspose.Cells that sets a worksheet to A2 paper size and outputs width and height in centimeters. | Explain how Aspose.Cells calculates PaperWidth and PaperHeight after assigning PaperSize. | Provide a C# snippet that checks whether the current page setup matches A2 dimensions and throws an exception if it does not.

using System;
using Aspose.Cells;

// Step‑by‑step example showing how to assign PaperSizeType.PaperA2 to a worksheet, retrieve PaperWidth and PaperHeight in inches, and save the workbook using Aspose.Cells for .NET.
class SetPaperSizeA2
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set the page setup paper size to A2
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA2;

        // Verify the updated dimensions (in inches)
        double widthInches = sheet.PageSetup.PaperWidth;
        double heightInches = sheet.PageSetup.PaperHeight;
        Console.WriteLine($"Paper size set to A2. Width = {widthInches} inches, Height = {heightInches} inches");

        // Save the workbook (optional verification step)
        workbook.Save("A2PaperSize.xlsx");
    }
}
