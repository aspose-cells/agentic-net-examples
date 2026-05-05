using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadFilterScenarios
{
    class Program
    {
        static void Main()
        {
            // Scenario 1: Load only workbook structure (sheet names, visibility, etc.)
            LoadIfExists("Sample1.xlsx", LoadStructureOnly);

            // Scenario 2: Load only cell values (skip formulas, formatting, drawings, etc.)
            LoadIfExists("Sample2.xlsx", LoadCellValuesOnly);

            // Scenario 3: Load a subset of sheets in a specific order.
            LoadIfExists("Sample3.xlsx", path => LoadSpecificSheets(path, new[] { 2, 0 }));

            // Scenario 4: Disable keeping unparsed data for read‑only operations.
            LoadIfExists("Sample4.xlsx", LoadWithoutUnparsedData);

            // Scenario 5: Ignore useless overlapping shapes.
            LoadIfExists("Sample5.xlsx", LoadIgnoringUselessShapes);
        }

        static void LoadIfExists(string filePath, Action<string> loadAction)
        {
            if (File.Exists(filePath))
            {
                loadAction(filePath);
            }
            else
            {
                Console.WriteLine($"\nFile not found: {filePath}. Skipping this scenario.");
            }
        }

        // Scenario 1 implementation
        static void LoadStructureOnly(string filePath)
        {
            var filter = new LoadFilter(LoadDataFilterOptions.Structure);
            var options = new LoadOptions { LoadFilter = filter };

            using var wb = new Workbook(filePath, options);

            Console.WriteLine("\nScenario 1 - Sheets loaded (structure only):");
            foreach (Worksheet ws in wb.Worksheets)
            {
                Console.WriteLine($"- {ws.Name} (Visible: {ws.IsVisible})");
                Console.WriteLine($"  Cells count: {ws.Cells.Count}");
            }
        }

        // Scenario 2 implementation
        static void LoadCellValuesOnly(string filePath)
        {
            var filter = new LoadFilter(LoadDataFilterOptions.CellValue);
            var options = new LoadOptions { LoadFilter = filter };

            using var wb = new Workbook(filePath, options);

            Console.WriteLine("\nScenario 2 - Cell values loaded:");
            Worksheet ws = wb.Worksheets[0];
            foreach (Cell cell in ws.Cells)
            {
                Console.WriteLine($"{cell.Name}: {cell.Value}");
            }
        }

        // Scenario 3 implementation
        static void LoadSpecificSheets(string filePath, int[] sheetOrder)
        {
            var filter = new CustomSheetOrderFilter(sheetOrder);
            var options = new LoadOptions { LoadFilter = filter };

            using var wb = new Workbook(filePath, options);

            Console.WriteLine("\nScenario 3 - Specific sheets loaded in custom order:");
            foreach (Worksheet ws in wb.Worksheets)
            {
                Console.WriteLine($"- Index: {ws.Index}, Name: {ws.Name}");
            }
        }

        // Scenario 4 implementation
        static void LoadWithoutUnparsedData(string filePath)
        {
            var options = new LoadOptions
            {
                KeepUnparsedData = false,
                LoadFilter = new LoadFilter()
            };

            using var wb = new Workbook(filePath, options);

            Console.WriteLine("\nScenario 4 - Workbook loaded without unparsed data.");
            Console.WriteLine($"Total worksheets: {wb.Worksheets.Count}");
        }

        // Scenario 5 implementation
        static void LoadIgnoringUselessShapes(string filePath)
        {
            var options = new LoadOptions
            {
                IgnoreUselessShapes = true,
                LoadFilter = new LoadFilter()
            };

            using var wb = new Workbook(filePath, options);

            Console.WriteLine("\nScenario 5 - Shapes after ignoring useless ones:");
            Console.WriteLine($"Shapes count in first sheet: {wb.Worksheets[0].Shapes.Count}");
        }

        // Custom LoadFilter to specify sheet loading order.
        class CustomSheetOrderFilter : LoadFilter
        {
            private readonly int[] _order;

            public CustomSheetOrderFilter(int[] order) : base(LoadDataFilterOptions.All)
            {
                _order = order;
            }

            public override int[] SheetsInLoadingOrder => _order;
        }
    }
}