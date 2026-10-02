// Title: Open a FODS workbook, change its default font to Arial 12 pt, and save it as an ODS file with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a .fods spreadsheet, updates the workbook’s default style to use Arial 12‑point font, and exports the result as an .ods file using Aspose.Cells. | Show how to programmatically set the default font for all cells in a workbook loaded from FODS and then save the workbook in ODS format with Aspose.Cells for .NET. | Provide a step‑by‑step C# example that changes the workbook‑wide default font and size, then converts the FODS document to ODS using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# change default font for entire workbook after loading FODS | Convert a .fods spreadsheet to .ods while applying Arial 12pt default style in .NET | How to set workbook default style font before saving as ODS with Aspose.Cells | C# code to load FODS, modify default workbook font, and export to ODS | Aspose.Cells default style modification example for FODS to ODS conversion
// Tags: set workbook-wide default font Aspose.Cells | import spreadsheet from FODS format C# | export workbook as ODS using Aspose.Cells | adjust default style for all cells .NET | apply Arial 12pt to workbook default style

using Aspose.Cells;
using System.Drawing;

// Loads a FODS workbook, changes the workbook’s default style to Arial 12 pt, and saves the result as an ODS file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the FODS workbook
        Workbook workbook = new Workbook("input.fods");

        // Change the default font for the entire workbook
        Style defaultStyle = workbook.DefaultStyle;
        defaultStyle.Font.Name = "Arial";   // Set desired font name
        defaultStyle.Font.Size = 12;        // Set desired font size
        workbook.DefaultStyle = defaultStyle;

        // Save the workbook as an ODS document
        workbook.Save("output.ods", SaveFormat.ODS);
    }
}
