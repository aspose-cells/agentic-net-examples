// Title: Delete blank rows and columns in a C# Aspose.Cells workbook before ODS export
// Description: Shows how to build a workbook, insert data with intentional empty rows and columns, remove those blanks using Cells.DeleteBlankRows() and Cells.DeleteBlankColumns(), and save the cleaned file as ODS with OdsSaveOptions.
// Keywords: Aspose.Cells DeleteBlankRows | Aspose.Cells DeleteBlankColumns | C# ODS export | remove empty rows Aspose | clean workbook before ODS | ODS conversion Aspose.Cells | C# spreadsheet cleanup
// Common Searches: C# Aspose.Cells delete blank rows | remove empty columns before ODS save | Aspose.Cells clean workbook ODS | DeleteBlankRows example C# | export cleaned spreadsheet to ODS Aspose
// Developer Intent: Remove all empty rows and columns from a workbook and produce a compact ODS document.
// Use Cases: Streamlining automatically generated reports that contain placeholder rows or columns prior to ODS distribution. | Preparing data extracts from databases where sparse rows/columns would inflate the ODS file size. | Batch processing multiple worksheets to eliminate blanks and deliver lean ODS files for downstream systems.
// AI Prompts: Write C# code that iterates through every worksheet in a workbook, deletes blank rows and columns with Aspose.Cells, and saves each sheet as an individual ODS file. | Explain how DeleteBlankRows and DeleteBlankColumns affect merged cells, formulas, and named ranges when cleaning a workbook for ODS export. | Provide a sample showing how to configure OdsSaveOptions (e.g., compression level, custom namespaces) after removing empty rows and columns.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Shows how to build a workbook, insert data with intentional empty rows and columns, remove those blanks using Cells.DeleteBlankRows() and Cells.DeleteBlankColumns(), and save the cleaned file as ODS with OdsSaveOptions.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add sample data with intentional blank rows and columns
        cells["A1"].PutValue("Header");
        cells["A2"].PutValue("Row1");
        // Row 3 is left blank
        cells["A4"].PutValue("Row2");

        // Column B is left blank
        cells["C1"].PutValue("ColCHeader");
        cells["C2"].PutValue("DataC1");
        cells["C4"].PutValue("DataC2");

        // Remove all blank rows and columns
        cells.DeleteBlankRows();
        cells.DeleteBlankColumns();

        // Save the cleaned workbook as ODS
        OdsSaveOptions odsOptions = new OdsSaveOptions();
        workbook.Save("CleanedWorkbook.ods", odsOptions);
    }
}
