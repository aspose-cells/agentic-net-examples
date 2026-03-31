using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsFilterDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create a new workbook --------------------
            Workbook workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];

            // -------------------- Populate sample data --------------------
            // Headers
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Product");
            ws.Cells["C1"].PutValue("Price");
            ws.Cells["D1"].PutValue("Date");
            ws.Cells["E1"].PutValue("Quantity");

            // Data rows
            ws.Cells["A2"].PutValue("Fruit");
            ws.Cells["B2"].PutValue("Apple");
            ws.Cells["C2"].PutValue(1.20);
            ws.Cells["D2"].PutValue(new DateTime(2022, 1, 15));
            ws.Cells["E2"].PutValue(50);

            ws.Cells["A3"].PutValue("Fruit");
            ws.Cells["B3"].PutValue("Banana");
            ws.Cells["C3"].PutValue(0.80);
            ws.Cells["D3"].PutValue(new DateTime(2022, 2, 10));
            ws.Cells["E3"].PutValue(30);

            ws.Cells["A4"].PutValue("Vegetable");
            ws.Cells["B4"].PutValue("Carrot");
            ws.Cells["C4"].PutValue(0.60);
            ws.Cells["D4"].PutValue(new DateTime(2021, 12, 5));
            ws.Cells["E4"].PutValue(40);

            ws.Cells["A5"].PutValue("Vegetable");
            ws.Cells["B5"].PutValue("Broccoli");
            ws.Cells["C5"].PutValue(1.00);
            ws.Cells["D5"].PutValue(new DateTime(2022, 3, 20));
            ws.Cells["E5"].PutValue(20);

            ws.Cells["A6"].PutValue("Fruit");
            ws.Cells["B6"].PutValue("Orange");
            ws.Cells["C6"].PutValue(1.10);
            ws.Cells["D6"].PutValue(new DateTime(2022, 1, 25));
            ws.Cells["E6"].PutValue(60);

            // -------------------- Apply AutoFilter to the whole table --------------------
            ws.AutoFilter.Range = "A1:E6";

            // -------------------- Simple filter: show only Fruit category --------------------
            ws.AutoFilter.AddFilter(0, "Fruit"); // fieldIndex 0 corresponds to column A (Category)
            ws.AutoFilter.Refresh();

            // -------------------- Custom filter: Price > 1.00 --------------------
            ws.AutoFilter.Custom(2, FilterOperatorType.GreaterThan, 1.00);
            ws.AutoFilter.Refresh();

            // -------------------- Date filter: Year = 2022 --------------------
            ws.AutoFilter.AddDateFilter(3, DateTimeGroupingType.Year, 2022, 0, 0, 0, 0, 0);
            ws.AutoFilter.Refresh();

            // -------------------- Fill color filter on Quantity column --------------------
            CellsColor fillFg = workbook.CreateCellsColor();
            fillFg.Color = Color.Yellow;
            CellsColor fillBg = workbook.CreateCellsColor();
            fillBg.Color = Color.White;
            ws.AutoFilter.AddFillColorFilter(4, BackgroundType.Solid, fillFg, fillBg);
            ws.AutoFilter.Refresh();

            // -------------------- Font color filter on Product column --------------------
            CellsColor fontColor = workbook.CreateCellsColor();
            fontColor.Color = Color.Red;
            ws.AutoFilter.AddFontColorFilter(1, fontColor);
            ws.AutoFilter.Refresh();

            // -------------------- Icon filter on Quantity column --------------------
            ws.AutoFilter.AddIconFilter(4, IconSetType.Arrows3, 1); // Show rows with first icon
            ws.AutoFilter.Refresh();

            // -------------------- Top 10 filter: top 3 quantities --------------------
            ws.AutoFilter.FilterTop10(4, true, false, 3);
            ws.AutoFilter.Refresh();

            // -------------------- Dynamic filter: Above Average on Price column --------------------
            ws.AutoFilter.DynamicFilter(2, DynamicFilterType.AboveAverage);
            ws.AutoFilter.Refresh();

            // -------------------- Match blanks on Category column --------------------
            ws.AutoFilter.MatchBlanks(0);
            ws.AutoFilter.Refresh();

            // -------------------- Match non-blanks on Category column (clears previous blank match) --------------------
            ws.AutoFilter.MatchNonBlanks(0);
            ws.AutoFilter.Refresh();

            // -------------------- Remove specific filter (Category = Fruit) --------------------
            ws.AutoFilter.RemoveFilter(0, "Fruit");
            ws.AutoFilter.Refresh();

            // -------------------- Remove all filters --------------------
            ws.AutoFilter.ShowAll();

            // -------------------- Advanced filter: copy unique rows where Category = Fruit --------------------
            // Prepare criteria range (A8:B9) where A8 is header, B8 is header, A9 contains "Fruit"
            ws.Cells["A8"].PutValue("Category");
            ws.Cells["B8"].PutValue("Product");
            ws.Cells["A9"].PutValue("Fruit");
            // Apply advanced filter: copy results to G1:H10
            ws.AdvancedFilter(false, "A2:E6", "A8:B9", "G1:H10", true);
            // Note: AdvancedFilter with isFilter = false copies the result; true would filter in place.

            // -------------------- Worksheet.Filter using CellArea (auto‑filter arrow) --------------------
            CellArea filterArea = new CellArea
            {
                StartRow = 0,      // Row 1 (zero‑based)
                StartColumn = 0,   // Column A
                EndRow = 5,        // Row 6
                EndColumn = 4      // Column E
            };
            ws.Filter(filterArea); // Applies an auto‑filter arrow to the defined area

            // -------------------- Save the workbook --------------------
            workbook.Save("FilteredWorkbook.xlsx");
        }
    }
}