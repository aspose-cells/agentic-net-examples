using System;
using Aspose.Cells;
using Aspose.Cells.Ods; // Required for OdsSaveOptions

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Sample data with intentional blank rows and columns
        cells["A1"].PutValue("Header");
        cells["A2"].PutValue("Data1");
        // Row 3 is left blank
        cells["A4"].PutValue("Data2"); // Blank row at A3 will be removed

        // Column B is left blank; data starts in column C
        cells["C1"].PutValue("Column C Header");
        cells["C2"].PutValue("Column C Data");

        // Remove all blank rows from the worksheet
        cells.DeleteBlankRows();

        // Remove all blank columns from the worksheet
        cells.DeleteBlankColumns();

        // Prepare ODS save options (optional settings can be adjusted here)
        OdsSaveOptions odsOptions = new OdsSaveOptions
        {
            // Example: ignore pivot tables when saving to ODS
            IgnorePivotTables = true
        };

        // Save the cleaned workbook as an ODS file
        workbook.Save("CleanedWorkbook.ods", odsOptions);
    }
}