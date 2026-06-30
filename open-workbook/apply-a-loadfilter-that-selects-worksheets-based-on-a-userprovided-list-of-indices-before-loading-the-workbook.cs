using System;
using Aspose.Cells;

namespace LoadFilterExample
{
    // Author: Aspose.Cells .NET example
    // Custom LoadFilter that loads only the worksheets whose indexes are provided by the user.
    public class IndexBasedLoadFilter : LoadFilter
    {
        private readonly int[] _sheetIndexes;

        // Constructor receives the list of worksheet indexes to load.
        public IndexBasedLoadFilter(int[] sheetIndexes)
        {
            _sheetIndexes = sheetIndexes ?? Array.Empty<int>();
        }

        // Override the SheetsInLoadingOrder property to return the user‑specified indexes.
        public override int[] SheetsInLoadingOrder => _sheetIndexes;

        // Optional: you can further control loading per sheet by overriding StartSheet.
        // Here we simply keep the default behavior.
        public override void StartSheet(Worksheet sheet)
        {
            base.StartSheet(sheet);
        }
    }

    class Program
    {
        static void Main()
        {
            // Example user‑provided list of worksheet indexes to load (zero‑based).
            int[] userSelectedIndexes = new int[] { 0, 2, 4 };

            // Prepare LoadOptions and assign the custom LoadFilter.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new IndexBasedLoadFilter(userSelectedIndexes);

            // Load the workbook using the specified LoadOptions.
            // Only the worksheets with indexes 0, 2, and 4 will be loaded.
            Workbook workbook = new Workbook("Template.xlsx", loadOptions);

            // Save the loaded workbook (or further process as needed).
            workbook.Save("FilteredOutput.xlsx");
        }
    }
}