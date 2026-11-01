// Title: How to clear a worksheet’s background image by setting BackgroundImage to null with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# clear worksheet background image | set worksheet.BackgroundImage = null to remove background picture in Excel | remove Excel sheet background image using Aspose.Cells .NET API | programmatically delete worksheet background picture in a .xlsx file | how to erase worksheet background image with C# Aspose.Cells
// Tags: set BackgroundImage to null Aspose.Cells | remove worksheet background picture .NET | save workbook after clearing background Aspose.Cells | background image deletion Excel C# | worksheet background image property manipulation Aspose.Cells

using Aspose.Cells;
using System.IO;

// The example loads an existing Excel workbook, accesses the first worksheet, clears its background image by assigning null to the BackgroundImage property, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (lifecycle rule: load)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the target worksheet (e.g., the first sheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Clear the worksheet's background image by assigning a null stream (feature rule)
        sheet.BackgroundImage = null;

        // Save the modified workbook (lifecycle rule: save)
        workbook.Save("output.xlsx");
    }
}
