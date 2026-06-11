using System;
using Aspose.Cells;

namespace LoadFirstWorksheetExample
{
    // Custom LoadFilter that loads only the first worksheet (index 0)
    class FirstSheetOnlyLoadFilter : LoadFilter
    {
        // Define the order of sheets to load – only sheet 0
        private readonly int[] _sheetsOrder = new int[] { 0 };

        // Constructor sets the default filter option to load all data for the selected sheet
        public FirstSheetOnlyLoadFilter() : base(LoadDataFilterOptions.All) { }

        // Override the read‑only property to return our custom sheet order
        public override int[] SheetsInLoadingOrder => _sheetsOrder;
    }

    class Program
    {
        static void Main()
        {
            // Path to the large source workbook
            string sourcePath = "largeWorkbook.xlsx";

            // Configure LoadOptions with the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new FirstSheetOnlyLoadFilter();

            // Load the workbook – only the first worksheet will be loaded
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save the result (contains only the first worksheet)
            string outputPath = "firstWorksheetOnly.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"First worksheet saved to: {outputPath}");
        }
    }
}