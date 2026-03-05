using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsFilterDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a workbook and populate sample data
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Category");
            cells["C1"].PutValue("Amount");
            cells["D1"].PutValue("Date");
            cells["E1"].PutValue("Status");

            // Sample rows
            for (int i = 2; i <= 11; i++)
            {
                cells[$"A{i}"].PutValue(i - 1); // ID
                cells[$"B{i}"].PutValue(i % 2 == 0 ? "Food" : "Drink"); // Category
                cells[$"C{i}"].PutValue(10 * i); // Amount
                cells[$"D{i}"].PutValue(new DateTime(2022, (i % 12) + 1, 1)); // Date
                cells[$"E{i}"].PutValue(i % 3 == 0 ? "Active" : "Inactive"); // Status
            }

            // -----------------------------------------------------------------
            // 2. Apply AutoFilter to the range A1:E11
            // -----------------------------------------------------------------
            sheet.AutoFilter.Range = "A1:E11";

            // 2.1 Custom numeric filter: Amount > 50
            sheet.AutoFilter.Custom(2, FilterOperatorType.GreaterThan, 50);
            sheet.AutoFilter.Refresh();

            // 2.2 Add a fill color filter on Category column (index 1)
            // Create a red fill color
            CellsColor redColor = workbook.CreateCellsColor();
            redColor.Color = Color.Red;
            // Apply the fill color filter (foreground red, background white)
            sheet.AutoFilter.AddFillColorFilter(1, BackgroundType.Solid, redColor, redColor);
            sheet.AutoFilter.Refresh();

            // 2.3 Add a date filter on Date column (index 3) for the year 2022
            sheet.AutoFilter.AddDateFilter(3, DateTimeGroupingType.Year, 2022, 0, 0, 0, 0, 0);
            sheet.AutoFilter.Refresh();

            // 2.4 Apply Top10 filter on Amount column (index 2) to show top 3 values
            sheet.AutoFilter.FilterTop10(2, false, false, 3);
            sheet.AutoFilter.Refresh();

            // 2.5 Remove the custom filter on Amount column
            sheet.AutoFilter.RemoveFilter(2);
            sheet.AutoFilter.Refresh();

            // -----------------------------------------------------------------
            // 3. Use AdvancedFilter to copy rows where Status = "Active"
            // -----------------------------------------------------------------
            // Define criteria range (G1:G2)
            cells["G1"].PutValue("Status");
            cells["G2"].PutValue("Active");
            // Apply advanced filter: copy matching rows to I1
            sheet.AdvancedFilter(false, "A1:E11", "G1:G2", "I1", false);
            // Note: The filtered data will be placed starting at I1

            // -----------------------------------------------------------------
            // 4. Save the workbook with all filters applied
            // -----------------------------------------------------------------
            string outputPath = "FilterDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");

            // -----------------------------------------------------------------
            // 5. Demonstrate LoadFilter: load only the workbook structure (no cell data)
            // -----------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions();
            // Load only the structure (sheets, formatting, etc.)
            loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.Structure);
            Workbook structureOnlyWb = new Workbook(outputPath, loadOptions);
            // Save the structure‑only workbook
            string structurePath = "FilterDemo_StructureOnly.xlsx";
            structureOnlyWb.Save(structurePath);
            Console.WriteLine($"Structure‑only workbook saved to {structurePath}");

            // -----------------------------------------------------------------
            // 6. Clean up
            // -----------------------------------------------------------------
            workbook.Dispose();
            structureOnlyWb.Dispose();
        }
    }
}