using System;
using System.Data;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // 1. Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells collection
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // ------------------------------------------------------------
        // 2. Create a named range using address strings (CreateRange overload)
        AsposeRange rangeA = cells.CreateRange("A1:C3");
        rangeA.Name = "DataBlock";

        // Fill the range with sample data
        for (int i = 0; i < rangeA.RowCount; i++)
        {
            for (int j = 0; j < rangeA.ColumnCount; j++)
            {
                rangeA[i, j].PutValue($"R{i + 1}C{j + 1}");
            }
        }

        // ------------------------------------------------------------
        // 3. Create another range using numeric indices (CreateRange overload)
        // Rows 5‑7 (zero‑based), columns A‑D (0‑3)
        AsposeRange rangeB = cells.CreateRange(5, 0, 3, 4);
        rangeB.Name = "CopySource";

        // Populate this range with numeric values
        for (int i = 0; i < rangeB.RowCount; i++)
        {
            for (int j = 0; j < rangeB.ColumnCount; j++)
            {
                rangeB[i, j].PutValue(i * 10 + j);
            }
        }

        // ------------------------------------------------------------
        // 4. Copy values from rangeB to a destination range using CopyValue
        AsposeRange destRange = cells.CreateRange(10, 0, rangeB.RowCount, rangeB.ColumnCount);
        destRange.CopyValue(rangeB);

        // ------------------------------------------------------------
        // 5. Add a range reference to the worksheet's Cells collection (AddRange)
        cells.AddRange(rangeA);

        // ------------------------------------------------------------
        // 6. Clear a sub‑range inside DataBlock (ClearRange overload)
        // Clears cells B1:C2 (rows 0‑1, columns 1‑2)
        cells.ClearRange(0, 1, 1, 2);

        // ------------------------------------------------------------
        // 7. Merge a block of cells and then unmerge it
        // Merge A13:B14 (rows 12‑13, columns 0‑1)
        cells.Merge(12, 0, 2, 2);
        // Create a Range object representing the merged area and unmerge it
        AsposeRange merged = cells.CreateRange(12, 0, 2, 2);
        merged.UnMerge();

        // ------------------------------------------------------------
        // 8. Group rows 0‑4 and columns 0‑2, then collapse and expand the groups
        cells.GroupRows(0, 4, true);
        cells.GroupColumns(0, 2, true);
        // Collapse the first group (index 0)
        cells.HideGroupDetail(true, 0);
        // Expand the first group again
        cells.ShowGroupDetail(true, 0);

        // ------------------------------------------------------------
        // 9. Apply a style to the entire copied range (ApplyStyle)
        Style style = workbook.CreateStyle();
        style.ForegroundColor = Color.LightYellow;
        style.Pattern = BackgroundType.Solid;
        style.Font.IsBold = true;

        destRange.ApplyStyle(style, new StyleFlag { FontBold = true, CellShading = true });

        // ------------------------------------------------------------
        // 10. Export the copied range to a DataTable (ExportDataTable)
        DataTable dt = destRange.ExportDataTable();

        // Write some information about the export back to the sheet
        cells["G1"].PutValue("Exported Rows");
        cells["H1"].PutValue(dt.Rows.Count);
        cells["G2"].PutValue("Exported Columns");
        cells["H2"].PutValue(dt.Columns.Count);

        // ------------------------------------------------------------
        // 11. Save the workbook (lifecycle: save)
        workbook.Save("AdvancedRangeDemo.xlsx");
    }
}