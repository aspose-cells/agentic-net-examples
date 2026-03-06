using System;
using Aspose.Cells;

namespace LoadFilterExamples
{
    class Program
    {
        static void Main()
        {
            // Scenario 1: Load only the workbook structure for all sheets except "DataSheet",
            // which is loaded with all data (cells, formulas, formatting, etc.).
            LoadOptions options1 = new LoadOptions();
            options1.LoadFilter = new CustomLoadFilterBySheet();
            Workbook wb1 = new Workbook("Template.xlsx", options1);
            wb1.Save("FilteredStructure.xlsx");

            // Scenario 2: Load only string and numeric cell values for every sheet.
            // This reduces memory usage when other data (charts, shapes, etc.) is not needed.
            LoadOptions options2 = new LoadOptions();
            options2.LoadFilter = new LoadFilter(LoadDataFilterOptions.CellString | LoadDataFilterOptions.CellNumeric);
            Workbook wb2 = new Workbook("Template.xlsx", options2);
            wb2.Save("FilteredCellValues.xlsx");

            // Scenario 3: Load sheets in a custom order and skip sheets not listed.
            // The array contains sheet indices as they appear in the source file.
            LoadOptions options3 = new LoadOptions();
            options3.LoadFilter = new CustomOrderedLoadFilter(new int[] { 2, 0 }); // Load Sheet3 first, then Sheet1
            Workbook wb3 = new Workbook("Template.xlsx", options3);
            wb3.Save("CustomOrder.xlsx");

            // Scenario 4: Optimize loading for large files by ignoring useless shapes
            // and disabling storage of unparsed data (e.g., comments, metadata).
            LoadOptions options4 = new LoadOptions();
            options4.IgnoreUselessShapes = true;
            options4.KeepUnparsedData = false;
            options4.LoadFilter = new LoadFilter(LoadDataFilterOptions.All);
            Workbook wb4 = new Workbook("Template.xlsx", options4);
            wb4.Save("OptimizedLoad.xlsx");
        }

        // Custom filter that loads full data only for a specific sheet.
        class CustomLoadFilterBySheet : LoadFilter
        {
            public override void StartSheet(Worksheet sheet)
            {
                // If the sheet name matches "DataSheet", load everything.
                // Otherwise, load only the workbook structure (no cell data, charts, etc.).
                if (sheet.Name.Equals("DataSheet", StringComparison.OrdinalIgnoreCase))
                {
                    LoadDataFilterOptions = LoadDataFilterOptions.All;
                }
                else
                {
                    LoadDataFilterOptions = LoadDataFilterOptions.Structure;
                }
            }
        }

        // Custom filter that defines a specific sheet loading order.
        class CustomOrderedLoadFilter : LoadFilter
        {
            private readonly int[] _order;

            public CustomOrderedLoadFilter(int[] order) : base(LoadDataFilterOptions.All)
            {
                _order = order;
            }

            // Override the read‑only property to supply the desired order.
            public override int[] SheetsInLoadingOrder => _order;
        }
    }
}