// Title: C# – Set Worksheet Left Margin to 0.5 in and Save as Excel 97‑2003 (.xls) using Aspose.Cells
// Description: Demonstrates how to create a workbook, set the left page margin to 0.5 inches via PageSetup.LeftMarginInch, and export the sheet to the legacy Excel 97‑2003 format with XlsSaveOptions in Aspose.Cells for .NET.
// Keywords: Aspose.Cells left margin | PageSetup.LeftMarginInch | C# export to xls | Excel 97-2003 save options | worksheet margin settings | Aspose.Cells .NET example
// Common Searches: Aspose.Cells set left margin inches C# | Save workbook as .xls using Aspose.Cells | How to change page margins before exporting to Excel 97‑2003 | C# code for worksheet margin configuration Aspose
// Developer Intent: Apply a 0.5‑inch left margin to a worksheet and generate an Excel 97‑2003 (.xls) file.
// Use Cases: Generating printable reports that require a specific left margin for legacy Excel users. | Automating batch conversion of modern workbooks to .xls while preserving margin layout. | Preparing documents for archival systems that only accept Excel 97‑2003 format.
// AI Prompts: Show C# code to set all page margins (top, bottom, left, right) in inches with Aspose.Cells before saving to .xls. | Provide an example that sets the left margin to 0.5 in for multiple worksheets and saves each as a separate .xls file. | Explain how to combine margin adjustments with page scaling when exporting a workbook to Excel 97‑2003 using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, set the left page margin to 0.5 inches via PageSetup.LeftMarginInch, and export the sheet to the legacy Excel 97‑2003 format with XlsSaveOptions in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data (optional)
        sheet.Cells["A1"].PutValue("Left margin set to 0.5 inches");

        // Set the left margin in inches
        sheet.PageSetup.LeftMarginInch = 0.5;

        // Create save options for Excel 97‑2003 format
        XlsSaveOptions saveOptions = new XlsSaveOptions();

        // Save the workbook as an .xls file
        workbook.Save("LeftMarginDemo.xls", saveOptions);
    }
}
