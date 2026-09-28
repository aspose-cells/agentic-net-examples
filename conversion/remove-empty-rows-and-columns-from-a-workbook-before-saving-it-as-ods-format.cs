// Title: Remove empty rows and columns from every worksheet and export the workbook to ODS with Aspose.Cells for .NET
// AI Prompts: Write a C# program using Aspose.Cells that loads an Excel file, removes all completely empty rows and columns from each worksheet, and saves the result as an ODS file. | Show how to iterate through all worksheets in a workbook and call DeleteBlankRows and DeleteBlankColumns before converting the workbook to ODS format. | Provide sample code that cleans an Excel workbook by deleting blank rows/columns and then uses SaveFormat.ODS to generate an OpenDocument Spreadsheet.
// Common Searches: Aspose.Cells delete empty rows and columns before saving as ODS | C# remove blank rows from all worksheets and convert to ODS using Aspose.Cells | how to clean Excel workbook with DeleteBlankRows DeleteBlankColumns and export to ODS | Aspose.Cells example for removing empty columns and saving as OpenDocument Spreadsheet
// Tags: delete blank rows Aspose.Cells | delete blank columns Aspose.Cells | export workbook to ODS Aspose.Cells | clean Excel worksheets before ODS conversion | Aspose.Cells DeleteBlankRows DeleteBlankColumns

using Aspose.Cells;
using System;

// The code loads an Excel workbook, iterates through each worksheet, deletes completely empty rows and columns using DeleteBlankRows and DeleteBlankColumns, and then saves the cleaned workbook as an ODS file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your source file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Remove all completely empty rows
            sheet.Cells.DeleteBlankRows();

            // Remove all completely empty columns
            sheet.Cells.DeleteBlankColumns();
        }

        // Save the cleaned workbook in ODS format
        workbook.Save("output.ods", SaveFormat.ODS);
    }
}
