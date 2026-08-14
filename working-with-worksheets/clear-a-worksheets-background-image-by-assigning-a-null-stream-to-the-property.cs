// Title: Clear a Worksheet Background Image in Aspose.Cells for .NET (C#)
// Description: Shows how to delete a worksheet's background picture by assigning null to the Worksheet.BackgroundImage property and then saving the workbook. Includes optional loading of an existing file.
// Keywords: Aspose.Cells | C# | clear worksheet background | remove background image | Worksheet.BackgroundImage | null stream | Excel background removal | Aspose.Cells .NET | clear background image code
// Common Searches: Aspose.Cells clear worksheet background image | How to remove background picture from Excel sheet using Aspose.Cells C# | Set Worksheet.BackgroundImage to null | Delete worksheet background in Aspose.Cells .NET | Remove background image from workbook programmatically
// Developer Intent: Remove the background picture from a worksheet using Aspose.Cells for .NET.
// Use Cases: Prepare a clean Excel template by stripping any default background images before populating data. | Batch‑process multiple worksheets to ensure a uniform, background‑free appearance in generated reports. | Reset a workbook's visual style after reusing a sheet that previously had a custom background.
// AI Prompts: Write C# code with Aspose.Cells that loops through all worksheets in a workbook and clears each background image. | Provide an example that loads an existing Excel file, removes the background image from the first sheet, and saves the updated file. | Explain the effect of setting Worksheet.BackgroundImage to null on the resulting .xlsx file when using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to delete a worksheet's background picture by assigning null to the Worksheet.BackgroundImage property and then saving the workbook. Includes optional loading of an existing file.
class ClearWorksheetBackgroundImage
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // new Workbook("input.xlsx") to load
        Worksheet worksheet = workbook.Worksheets[0];

        // Example: set a background image (optional)
        // byte[] imgData = System.IO.File.ReadAllBytes("background.jpg");
        // worksheet.BackgroundImage = imgData;

        // Clear the background image by assigning null
        worksheet.BackgroundImage = null;

        // Save the workbook with the cleared background
        workbook.Save("ClearedBackground.xlsx");
    }
}
