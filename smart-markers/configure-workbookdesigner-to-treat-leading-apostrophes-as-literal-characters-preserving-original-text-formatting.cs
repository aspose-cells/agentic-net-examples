// Title: How to keep a leading apostrophe as literal text when using Aspose.Cells WorkbookDesigner in C#
// AI Prompts: Generate C# code that sets Workbook.Settings.QuotePrefixToStyle to false so a leading apostrophe remains part of the cell value after WorkbookDesigner processing. | Show the complete steps to create a workbook, insert a value with a leading apostrophe, disable quote‑prefix conversion, run WorkbookDesigner, and save the file using Aspose.Cells. | Explain how to prevent Aspose.Cells from treating a leading apostrophe as a formatting flag when working with smart markers in C#.
// Common Searches: Aspose.Cells C# preserve leading apostrophe in Excel cell | WorkbookDesigner keep apostrophe character in cell value | how to disable QuotePrefixToStyle setting in Aspose.Cells | prevent Aspose.Cells from stripping leading apostrophe with smart markers | save Excel file with literal apostrophe using Aspose.Cells C#
// Tags: Workbook.Settings.QuotePrefixToStyle configuration | apostrophe preservation Aspose.Cells | WorkbookDesigner literal apostrophe handling | disable quote prefix conversion C# | smart markers apostrophe handling

using System;
using Aspose.Cells;

// Sets Workbook.Settings.QuotePrefixToStyle to false so that a leading apostrophe in a cell remains part of the text, processes the workbook with WorkbookDesigner, and saves the result to an Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook (could also be loaded from a template file)
        Workbook workbook = new Workbook();

        // Configure the workbook so that leading apostrophes are treated as literal characters
        // Setting QuotePrefixToStyle to false prevents Aspose.Cells from converting the leading
        // apostrophe into a formatting flag (QuotePrefix). The apostrophe remains part of the cell value.
        workbook.Settings.QuotePrefixToStyle = false;

        // Example cell containing a leading apostrophe
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("'Sample text with leading apostrophe");

        // Initialize WorkbookDesigner with the configured workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Process smart markers if any (not required for this demonstration)
        designer.Process();

        // Save the result; the leading apostrophe is preserved as part of the cell's text
        workbook.Save("Result.xlsx");
    }
}
