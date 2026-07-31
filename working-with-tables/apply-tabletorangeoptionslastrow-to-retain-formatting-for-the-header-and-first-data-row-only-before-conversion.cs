// Title: Aspose.Cells C# – Convert ListObject to Range while keeping header and first row styles (TableToRangeOptions.LastRow)
// Description: Demonstrates how to create a workbook, style the header and the first data row, add a ListObject, and convert the table to a range using TableToRangeOptions with LastRow = 1 so that only those two rows retain their formatting.
// Keywords: Aspose.Cells TableToRangeOptions | C# convert ListObject to range | preserve header formatting Aspose | retain first data row style | .NET Excel table to range | LastRow option example
// Common Searches: Aspose.Cells keep header style when converting table to range | TableToRangeOptions LastRow C# example | convert Excel ListObject to range preserving first rows | Aspose.Cells table conversion options
// Developer Intent: Keep the header and first data row formatting while flattening a table into a range.
// Use Cases: Generate a report where only the top rows of a table need to stay visually styled after conversion. | Export a subset of a table for downstream processing while preserving its header and first row appearance. | Create a styled worksheet, apply a table, then flatten it without losing the initial row styles.
// AI Prompts: Show C# code using Aspose.Cells to convert a ListObject to a range and retain header and first row formatting with TableToRangeOptions.LastRow. | Explain how TableToRangeOptions.LastRow works in Aspose.Cells and give a practical example. | Provide a step‑by‑step guide to style the header and first data row of an Excel table and then flatten the table while keeping those styles.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableToRangeDemo
{
    // Demonstrates how to create a workbook, style the header and the first data row, add a ListObject, and convert the table to a range using TableToRangeOptions with LastRow = 1 so that only those two rows retain their formatting.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate header row (row 0)
            cells[0, 0].PutValue("ID");
            cells[0, 1].PutValue("Name");
            cells[0, 2].PutValue("Score");

            // Populate data rows (rows 1 to 5)
            for (int row = 1; row <= 5; row++)
            {
                cells[row, 0].PutValue(row);                     // ID
                cells[row, 1].PutValue($"User{row}");            // Name
                cells[row, 2].PutValue(50 + row * 5);            // Score
            }

            // Apply formatting to header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;
            StyleFlag headerFlag = new StyleFlag { FontBold = true, CellShading = true };
            cells.CreateRange(0, 0, 1, 3).ApplyStyle(headerStyle, headerFlag);

            // Apply formatting to the first data row (row 1)
            Style firstDataStyle = workbook.CreateStyle();
            firstDataStyle.ForegroundColor = System.Drawing.Color.LightYellow;
            firstDataStyle.Pattern = BackgroundType.Solid;
            StyleFlag firstDataFlag = new StyleFlag { CellShading = true };
            cells.CreateRange(1, 0, 1, 3).ApplyStyle(firstDataStyle, firstDataFlag);

            // Add a ListObject (table) that covers all rows (0‑5) and columns (0‑2)
            int tableIndex = sheet.ListObjects.Add(0, 0, 5, 2, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Convert the table to a range, retaining only header and first data row
            TableToRangeOptions options = new TableToRangeOptions
            {
                // LastRow is zero‑based; setting it to 1 keeps rows 0 and 1.
                LastRow = 1
            };
            table.ConvertToRange(options);

            // Save the workbook
            workbook.Save("TableToRange_WithHeaderAndFirstRow.xlsx");
        }
    }
}
