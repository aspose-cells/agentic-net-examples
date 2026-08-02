// Title: Freeze the first column and export to Excel 97‑2003 .xls with Aspose.Cells for .NET
// Description: Demonstrates creating a Workbook in C#, adding sample data, applying FreezePanes at cell B1 to lock column A, configuring XlsSaveOptions for the 97‑2003 format, and saving the file as FrozenColumns.xls for legacy Excel compatibility.
// Keywords: Aspose.Cells | C# | FreezePanes | freeze column | XlsSaveOptions | legacy Excel | Excel 97-2003 | save as .xls | frozen panes export | Aspose.Cells .NET
// Common Searches: Aspose.Cells freeze column C# example | Save workbook as .xls with frozen panes using Aspose.Cells | How to use XlsSaveOptions for legacy Excel format | Freeze first column before exporting to Excel 97-2003 | C# code to apply FreezePanes and save as .xls
// Developer Intent: Need to lock a worksheet column and generate an .xls file that works with older Excel versions.
// Use Cases: Generating client‑facing reports where the ID column must stay visible while scrolling, and the recipient uses Excel 2003. | Automating data extracts for legacy ERP systems that only accept .xls files with frozen panes. | Creating email‑ready spreadsheets with frozen header columns for quick review on older Office installations.
// AI Prompts: Provide C# code to freeze multiple columns and save the workbook as .xls with Aspose.Cells. | Show how to preserve frozen panes when converting an .xlsx file to .xls using Aspose.Cells. | Explain the XlsSaveOptions settings required to keep FreezePanes during export.

using System;
using Aspose.Cells;

// Demonstrates creating a Workbook in C#, adding sample data, applying FreezePanes at cell B1 to lock column A, configuring XlsSaveOptions for the 97‑2003 format, and saving the file as FrozenColumns.xls for legacy Excel compatibility.
class FreezeColumnsToXls
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["C1"].PutValue("Score");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");
        sheet.Cells["C2"].PutValue(85);
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Bob");
        sheet.Cells["C3"].PutValue(92);

        // Freeze the first column (column A)
        // Freeze panes at cell B1, with 0 frozen rows and 1 frozen column
        sheet.FreezePanes("B1", 0, 1);

        // Create XLS save options for legacy compatibility
        XlsSaveOptions saveOptions = new XlsSaveOptions();

        // Save the workbook as an Excel 97-2003 .xls file
        workbook.Save("FrozenColumns.xls", saveOptions);
    }
}
