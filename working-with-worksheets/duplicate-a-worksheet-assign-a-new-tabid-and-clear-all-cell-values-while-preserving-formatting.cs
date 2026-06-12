using System;
using Aspose.Cells;

namespace WorksheetDuplicateExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the source worksheet (first sheet by default)
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Original";
            // Populate some sample data with formatting to demonstrate preservation
            sourceSheet.Cells["A1"].PutValue("Header");
            sourceSheet.Cells["A1"].GetStyle().Font.IsBold = true;
            sourceSheet.Cells["A2"].PutValue(123);
            sourceSheet.Cells["A2"].GetStyle().ForegroundColor = System.Drawing.Color.LightYellow;
            sourceSheet.Cells["A2"].GetStyle().Pattern = BackgroundType.Solid;

            // Duplicate the worksheet using AddCopy (copies contents and formats)
            int copiedIndex = workbook.Worksheets.AddCopy(0);
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
            copiedSheet.Name = "Copy";

            // Assign a new TabId to the duplicated sheet
            // Here we simply set it to a value different from the source sheet
            copiedSheet.TabId = sourceSheet.TabId + 1000;

            // Clear all cell values while preserving formatting
            // ClearContents removes only the data, leaving styles intact
            int maxRow = copiedSheet.Cells.MaxRow;
            int maxColumn = copiedSheet.Cells.MaxColumn;
            // Ensure we cover the entire used range
            copiedSheet.Cells.ClearContents(0, 0, maxRow + 1, maxColumn + 1);

            // Save the workbook to verify the result
            workbook.Save("DuplicatedWorksheet.xlsx");
        }
    }
}