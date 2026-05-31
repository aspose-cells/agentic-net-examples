using System;
using Aspose.Cells;

namespace AsposeCellsLoadFilterExample
{
    // Custom LoadFilter that returns a specific sheet loading order
    public class CustomLoadFilter : LoadFilter
    {
        private readonly int[] _sheetsOrder;

        // Accept the desired sheet indices via constructor
        public CustomLoadFilter(int[] sheetsOrder) : base(LoadDataFilterOptions.All)
        {
            _sheetsOrder = sheetsOrder;
        }

        // Override the read‑only property to supply the indices
        public override int[] SheetsInLoadingOrder => _sheetsOrder;
    }

    class Program
    {
        static void Main()
        {
            // User‑provided list of sheet indexes to load (e.g., load sheet 0 and sheet 2)
            int[] selectedSheetIndexes = new int[] { 0, 2 };

            // Create the custom filter with the desired order
            LoadFilter filter = new CustomLoadFilter(selectedSheetIndexes);

            // Configure load options to use the filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = filter;

            // Load the workbook using the filter – only the specified sheets will be loaded
            Workbook workbook = new Workbook("Template.xlsx", loadOptions);

            // Demonstrate which sheets were loaded
            Console.WriteLine("Loaded worksheets count: " + workbook.Worksheets.Count);
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Console.WriteLine($"- Index: {ws.Index}, Name: {ws.Name}");
            }

            // Save the workbook (optional, demonstrates that saving works with filtered sheets)
            workbook.Save("FilteredOutput.xlsx");
        }
    }
}