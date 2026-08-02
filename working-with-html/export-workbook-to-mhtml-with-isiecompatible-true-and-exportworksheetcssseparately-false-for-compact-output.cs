// Title: Export Aspose.Cells Workbook to MHTML with IE Compatibility and Inline CSS (C#)
// Description: Creates a workbook, adds sample data, configures HtmlSaveOptions with IsIECompatible=true and ExportWorksheetCSSSeparately=false, and saves the result as a single compact MHTML file.
// Keywords: Aspose.Cells | C# | MHTML export | IsIECompatible | inline CSS | HtmlSaveOptions | compact output
// Common Searches: Aspose.Cells save as MHTML with IE mode | Export worksheet CSS inline Aspose.Cells | C# generate single‑file MHTML from workbook | HtmlSaveOptions ExportWorksheetCSSSeparately example
// Developer Intent: Produce a self‑contained MHTML document from a workbook that works in legacy Internet Explorer and avoids external CSS files.
// Use Cases: Email‑ready reports that render correctly in older IE browsers | Embedding spreadsheet content in a web page without additional style resources | Creating a portable, offline‑viewable archive of spreadsheet data
// AI Prompts: Show how to export a workbook to MHTML with IsIECompatible=true and ExportWorksheetCSSSeparately=false using Aspose.Cells for .NET. | Provide a C# snippet that saves multiple worksheets into one MHTML file with inline CSS and IE compatibility. | Explain the impact of IsIECompatible and ExportWorksheetCSSSeparately on the size and rendering of the generated MHTML.

using System;
using Aspose.Cells;

// Creates a workbook, adds sample data, configures HtmlSaveOptions with IsIECompatible=true and ExportWorksheetCSSSeparately=false, and saves the result as a single compact MHTML file.
class ExportWorkbookToMhtml
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Exported to MHTML with IE compatibility");

        // Configure HTML save options for MHTML output
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        // Enable IE compatibility mode
        saveOptions.IsIECompatible = true;
        // Export worksheet CSS inline (do not create separate CSS files) for a more compact file
        saveOptions.ExportWorksheetCSSSeparately = false;

        // Save the workbook as an MHTML file using the configured options
        workbook.Save("ExportedDocument.mht", saveOptions);
    }
}
