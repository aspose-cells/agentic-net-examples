// Title: Set ODS Worksheet Page Background to Light Gray with Aspose.Cells for .NET
// Description: Creates a new workbook, accesses the first worksheet, configures the ODS page background to LightGray via OdsPageBackground, and saves the file as an .ods document.
// Keywords: Aspose.Cells | C# | ODS | worksheet background | light gray | OdsPageBackground | page color | print contrast
// Common Searches: Aspose.Cells set ODS page background color C# | how to change worksheet background in ODS using .NET | light gray page background Aspose.Cells | set ODS page background before saving workbook | C# print page background color Aspose.Cells
// Developer Intent: Programmatically apply a light gray background to the ODS page layout of a worksheet.
// Use Cases: Enhance readability of printed ODS reports by adding a subtle gray page background. | Enforce a consistent visual style across automatically generated ODS files. | Meet corporate branding requirements that specify a particular page shade for exported spreadsheets.
// AI Prompts: Show how to set a custom RGB background color for an ODS worksheet with Aspose.Cells in C#. | Provide code to switch the ODS page background between a solid color and an image at runtime. | Explain methods to verify the background color after saving the workbook as an ODS file.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Creates a new workbook, accesses the first worksheet, configures the ODS page background to LightGray via OdsPageBackground, and saves the file as an .ods document.
class SetWorksheetBackground
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Configure the ODS page background to a light gray color
        OdsPageBackground background = sheet.PageSetup.ODSPageBackground;
        background.Type = OdsPageBackgroundType.Color;
        background.Color = Color.LightGray;

        // Save the workbook (ODS format supports the background color setting)
        workbook.Save("WorksheetWithBackground.ods");
    }
}
