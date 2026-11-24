// Title: How to clear a worksheet’s background image in Aspose.Cells for .NET by setting BackgroundImage to null
// Common Searches: Aspose.Cells C# remove worksheet background picture programmatically | clear Excel sheet background image using Aspose.Cells .NET | set worksheet BackgroundImage to null in Aspose.Cells example | delete worksheet background image before saving workbook with Aspose.Cells | remove background image from existing .xlsx file using Aspose.Cells API
// Tags: Aspose.Cells clear worksheet background image | BackgroundImage null assignment C# | clear worksheet background picture Aspose.Cells | reset worksheet background Aspose.Cells | programmatic removal of worksheet background image

using System;
using Aspose.Cells;

// The example loads an existing workbook, accesses a worksheet, clears its background image by assigning null to the BackgroundImage property, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any specific worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Clear the worksheet's background image by assigning null to the BackgroundImage property
        worksheet.BackgroundImage = null;

        // Save the workbook with the changes
        workbook.Save("output.xlsx");
    }
}
