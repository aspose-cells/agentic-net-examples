// Title: Clone a worksheet in Aspose.Cells for .NET – retain background image and page‑setup settings
// Description: Shows how to assign a background image and page‑setup (paper size, orientation, print area) to a worksheet, duplicate it with AddCopy, copy the page‑setup via PageSetup.Copy, transfer the background image byte stream, and save the result.
// Keywords: Aspose.Cells | C# | clone worksheet | AddCopy | background image | page setup copy | preserve print area | worksheet duplication | CopyOptions | Excel automation
// Common Searches: Aspose.Cells clone worksheet with background image | Copy page setup between worksheets .NET | Preserve background image when duplicating Excel sheet | AddCopy keep print area Aspose.Cells | Duplicate worksheet C# Aspose.Cells
// Developer Intent: Duplicate an existing worksheet while keeping its background image and all page‑setup properties exactly as they were.
// Use Cases: Create a master template sheet that includes a logo background and A4 landscape layout, then clone it for each monthly report without re‑applying formatting. | Generate localized versions of a financial model by cloning the original sheet, swapping only the data, and retaining the original print area and background. | Automate workbook generation where dozens of sheets share identical page‑setup and background graphics, ensuring consistency with a single AddCopy operation.
// AI Prompts: Write C# code using Aspose.Cells that clones a worksheet, copies its background image byte array, and duplicates all page‑setup settings, handling the case where the source sheet has no image. | Explain the role of CopyOptions in PageSetup.Copy for Aspose.Cells and give best‑practice tips for cloning worksheets with complex printing configurations. | Create a C# unit test that verifies the cloned worksheet contains the same background image bytes and identical PageSetup properties (paper size, orientation, print area) as the source sheet.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to assign a background image and page‑setup (paper size, orientation, print area) to a worksheet, duplicate it with AddCopy, copy the page‑setup via PageSetup.Copy, transfer the background image byte stream, and save the result.
class CloneWorksheetDemo
{
    static void Main()
    {
        // Create a source workbook and get its first worksheet
        Workbook srcWorkbook = new Workbook();
        Worksheet srcSheet = srcWorkbook.Worksheets[0];
        srcSheet.Name = "Original";

        // Set a background image for the source worksheet (if the file exists)
        string imagePath = "background.jpg";
        if (File.Exists(imagePath))
        {
            srcSheet.BackgroundImage = File.ReadAllBytes(imagePath);
        }

        // Configure page‑setup settings on the source worksheet
        srcSheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
        srcSheet.PageSetup.Orientation = PageOrientationType.Landscape;
        srcSheet.PageSetup.PrintArea = "A1:D20";

        // Clone the worksheet using AddCopy (creates a new sheet with copied contents)
        int clonedIndex = srcWorkbook.Worksheets.AddCopy(srcSheet.Index);
        Worksheet clonedSheet = srcWorkbook.Worksheets[clonedIndex];
        clonedSheet.Name = "Cloned";

        // Preserve the original page‑setup configuration
        clonedSheet.PageSetup.Copy(srcSheet.PageSetup, new CopyOptions());

        // Preserve the background image stream
        clonedSheet.BackgroundImage = srcSheet.BackgroundImage;

        // Save the workbook containing the cloned worksheet
        srcWorkbook.Save("ClonedWorksheet.xlsx");
    }
}
