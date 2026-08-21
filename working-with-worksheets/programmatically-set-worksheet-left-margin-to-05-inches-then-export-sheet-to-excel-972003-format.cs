// Title: C# – Set Worksheet Left Margin to 0.5 in and Save as Excel 97‑2003 (.xls) with Aspose.Cells
// Description: Creates a new workbook, sets the first worksheet's left page margin to 0.5 inches via PageSetup.LeftMarginInch, optionally adds sample data, configures XlsSaveOptions, and saves the file as an Excel 97‑2003 (.xls) document.
// Keywords: Aspose.Cells left margin C# | PageSetup.LeftMarginInch example | export to Excel 97-2003 | save workbook as .xls | legacy Excel format Aspose.Cells | C# worksheet margin settings
// Common Searches: how to set left margin 0.5 inches Aspose.Cells | C# save workbook as Excel 97-2003 file | Aspose.Cells PageSetup margin properties | export worksheet with custom margins to .xls | Aspose.Cells legacy Excel format example
// Developer Intent: Set a worksheet's left margin to 0.5 inches and export the workbook as an Excel 97‑2003 (.xls) file using Aspose.Cells for .NET.
// Use Cases: Produce printable reports with a precise left margin while delivering them in a legacy .xls format. | Automate batch generation of Excel 97‑2003 files where specific page margins are required for compliance. | Archive spreadsheets with predefined margins to guarantee consistent layout when opened in older Excel versions.
// AI Prompts: Generate C# code with Aspose.Cells that sets a worksheet's left margin to 0.5 inches and saves the workbook as an Excel 97‑2003 (.xls) file. | Show how to adjust all page margins (top, bottom, left, right) in Aspose.Cells and export the result to .xls using C#. | Explain the behavior of PageSetup.LeftMarginInch and how it interacts with XlsSaveOptions for creating legacy Excel files.

using System;
using Aspose.Cells;

// Creates a new workbook, sets the first worksheet's left page margin to 0.5 inches via PageSetup.LeftMarginInch, optionally adds sample data, configures XlsSaveOptions, and saves the file as an Excel 97‑2003 (.xls) document.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set the left margin to 0.5 inches
        sheet.PageSetup.LeftMarginInch = 0.5;

        // Add sample data (optional, just to have some content)
        sheet.Cells["A1"].PutValue("Left margin set to 0.5 inches");

        // Create save options for Excel 97‑2003 format
        XlsSaveOptions saveOptions = new XlsSaveOptions();

        // Save the workbook as an .xls file using the specified options
        workbook.Save("LeftMarginDemo.xls", saveOptions);
    }
}
