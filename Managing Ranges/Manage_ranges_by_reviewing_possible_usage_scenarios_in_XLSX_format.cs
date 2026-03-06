using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Create a range using address strings
            AsposeRange dataRange = cells.CreateRange("A1", "B3");
            dataRange.Name = "SampleData";

            dataRange[0, 0].PutValue("Product");
            dataRange[0, 1].PutValue("Price");
            dataRange[1, 0].PutValue("Laptop");
            dataRange[1, 1].PutValue(1200.5);
            dataRange[2, 0].PutValue("Phone");
            dataRange[2, 1].PutValue(799.99);

            // 3. Create another range using address strings (destination starts at A5)
            AsposeRange destRange = cells.CreateRange("A5", "B7");
            destRange.Name = "CopiedData";
            destRange.CopyValue(dataRange);

            // 4. Clear contents of a sub‑range (C2:C3)
            AsposeRange clearRange = cells.CreateRange("C2", "C3");
            clearRange[0, 0].PutValue("Temp1");
            clearRange[1, 0].PutValue("Temp2");
            clearRange.ClearContents();

            // 5. Merge a range (E1:G2) and set a value
            AsposeRange mergeRange = cells.CreateRange("E1", "G2");
            mergeRange.Merge();
            mergeRange[0, 0].PutValue("Merged Cell");

            Style mergedStyle = workbook.CreateStyle();
            mergedStyle.Font.Color = Color.White;
            mergedStyle.Font.IsBold = true;
            mergedStyle.ForegroundColor = Color.DarkBlue;
            mergedStyle.Pattern = BackgroundType.Solid;

            mergeRange.ApplyStyle(mergedStyle, new StyleFlag
            {
                Font = true,
                CellShading = true
            });

            // 6. Load an existing workbook and work with its ranges
            string templatePath = "Template.xlsx";
            if (System.IO.File.Exists(templatePath))
            {
                Workbook templateWb = new Workbook(templatePath);
                Worksheet tmplSheet = templateWb.Worksheets[0];
                Cells tmplCells = tmplSheet.Cells;

                int maxRow = tmplCells.MaxDataRow;
                int maxCol = tmplCells.MaxDataColumn;

                if (maxRow >= 0 && maxCol >= 0)
                {
                    // Create range covering used area (address strings)
                    string startAddr = "A1";
                    string endAddr = CellsHelper.CellIndexToName(maxRow, maxCol);
                    AsposeRange templateRange = tmplCells.CreateRange(startAddr, endAddr);
                    templateRange.Name = "TemplateArea";

                    // Destination starts at A10
                    string destStart = "A10";
                    int destStartRow, destStartCol;
                    CellsHelper.CellNameToIndex(destStart, out destStartRow, out destStartCol);

                    int destEndRow = destStartRow + templateRange.RowCount - 1;
                    int destEndCol = destStartCol + templateRange.ColumnCount - 1;
                    string destEnd = CellsHelper.CellIndexToName(destEndRow, destEndCol);

                    AsposeRange targetRange = cells.CreateRange(destStart, destEnd);
                    targetRange.Copy(templateRange);
                }
            }

            // 7. Save the workbook
            string outputPath = "ManagedRangesDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}