using System;
using Aspose.Cells;

namespace AsposeCellsIndentExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data representing a simple hierarchical list in column A
                cells["A1"].PutValue("Item 1");
                cells["A2"].PutValue("Subitem 1.1");
                cells["A3"].PutValue("Subitem 1.2");
                cells["A4"].PutValue("Item 2");
                cells["A5"].PutValue("Subitem 2.1");

                // -------------------------------------------------
                // Create a style that sets indentation to two levels
                // -------------------------------------------------
                Style indentStyle = workbook.CreateStyle();
                indentStyle.IndentLevel = 2; // two indent levels
                indentStyle.HorizontalAlignment = TextAlignmentType.Left; // required for indent to work

                // Indicate that only the indent property should be applied
                StyleFlag flag = new StyleFlag();
                flag.Indent = true;

                // -------------------------------------------------
                // Define the range that corresponds to the nested list
                // -------------------------------------------------
                // Rows 1‑5 in column A (zero‑based indices 0‑4, column 0)
                Aspose.Cells.Range listRange = cells.CreateRange(0, 0, 5, 1);

                // Apply the indentation style to the entire range
                listRange.ApplyStyle(indentStyle, flag);

                // Save the workbook
                workbook.Save("NestedListIndent.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                // Log or display the error details
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}