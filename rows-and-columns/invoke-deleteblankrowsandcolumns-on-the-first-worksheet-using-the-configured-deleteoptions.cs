using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (contains one worksheet by default)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data with intentional blank rows and columns
        cells["A1"].PutValue("Header");
        cells["A2"].PutValue("Data1");
        // Row 3 is left blank
        cells["A4"].PutValue("Data2"); // This creates a blank row at index 2 (row 3)

        // Column B is left blank; data starts in column C
        cells["C1"].PutValue("ColC");
        cells["C2"].PutValue("ValC2");

        // Configure DeleteOptions (e.g., update references after deletion)
        DeleteOptions options = new DeleteOptions
        {
            UpdateReference = true
        };

        // Delete blank rows and columns using the configured options
        cells.DeleteBlankRows(options);
        cells.DeleteBlankColumns(options);

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}