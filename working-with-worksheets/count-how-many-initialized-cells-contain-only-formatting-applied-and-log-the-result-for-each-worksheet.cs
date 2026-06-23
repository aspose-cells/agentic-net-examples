using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class CountFormattingOnlyCells
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // Replace with new Workbook("input.xlsx") to load

            // -------------------------------------------------
            // Sample data: add some cells with only formatting
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            Cells cells1 = sheet1.Cells;

            // Cell with value only
            cells1["A1"].PutValue("Data");

            // Cell with formatting only
            Style fmtOnly = workbook.CreateStyle();
            fmtOnly.Font.Color = System.Drawing.Color.Red;
            cells1["B2"].SetStyle(fmtOnly);

            // Cell with both value and formatting
            cells1["C3"].PutValue(123);
            cells1["C3"].SetStyle(fmtOnly);

            // -------------------------------------------------
            // Count cells that have formatting but no value
            // -------------------------------------------------
            foreach (Worksheet ws in workbook.Worksheets)
            {
                int formattingOnlyCount = 0;

                // Iterate over all instantiated cells in the worksheet
                foreach (Cell cell in ws.Cells)
                {
                    // Check if the cell has no value (null) and has a style different from the default style
                    if (cell.Value == null && !cell.GetStyle().Equals(workbook.DefaultStyle))
                    {
                        formattingOnlyCount++;
                    }
                }

                Console.WriteLine($"Worksheet \"{ws.Name}\": {formattingOnlyCount} cells contain only formatting.");
            }

            // Save the workbook if needed
            workbook.Save("output.xlsx");
        }
    }
}