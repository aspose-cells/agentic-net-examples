// Title: C# – Load selected worksheets by index using a custom Aspose.Cells LoadFilter
// Description: Demonstrates how to create a CustomLoadFilter that accepts an array of zero‑based worksheet indexes, assigns it to LoadOptions, and loads only those sheets when opening a Workbook. The example shows enumeration of the loaded sheets and optional saving, helping reduce memory usage and preserve user‑defined sheet order.
// Keywords: Aspose.Cells | LoadFilter | C# | load specific worksheets | worksheet index | LoadOptions example | select sheets on workbook load | Excel memory optimization | GitHub Aspose.Cells sample | custom sheet order
// Common Searches: Aspose.Cells load only certain sheets C# | Custom LoadFilter example for selective worksheet loading | How to load worksheets by index with Aspose.Cells | LoadOptions SheetsInLoadingOrder property usage | Reduce memory when opening large Excel files Aspose
// Developer Intent: Open a workbook while loading only the worksheets whose indexes are supplied by the user.
// Use Cases: Extract data from a large template by loading just the required sheets, minimizing RAM consumption. | Maintain a user‑defined sheet sequence when opening a workbook for further processing. | Speed up report generation by pre‑filtering worksheets before any data extraction.
// AI Prompts: Generate a C# snippet that uses Aspose.Cells LoadFilter to load worksheets 2, 4, and 5 and then saves the filtered workbook. | Explain how to adapt the CustomLoadFilter to exclude sheets based on their names instead of indexes. | Provide error‑handling code for cases where user‑provided sheet indexes are out of range when using a custom LoadFilter.

using System;
using Aspose.Cells;

namespace LoadFilterExample
{
    // Custom LoadFilter that specifies which sheets to load and in what order
    // Demonstrates how to create a CustomLoadFilter that accepts an array of zero‑based worksheet indexes, assigns it to LoadOptions, and loads only those sheets when opening a Workbook. The example shows enumeration of the loaded sheets and optional saving, helping reduce memory usage and preserve user‑defined sheet order.
    public class CustomLoadFilter : LoadFilter
    {
        private readonly int[] _sheetsOrder;

        // Constructor accepts an array of sheet indexes to load
        public CustomLoadFilter(int[] sheetsOrder) : base(LoadDataFilterOptions.All)
        {
            _sheetsOrder = sheetsOrder;
        }

        // Override the read‑only property to return the desired sheet order
        public override int[] SheetsInLoadingOrder => _sheetsOrder;
    }

    class Program
    {
        static void Main()
        {
            // User‑provided list of sheet indexes to load (zero‑based)
            int[] selectedSheetIndexes = new int[] { 0, 2 }; // example: load first and third sheets

            // Create LoadOptions and assign the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new CustomLoadFilter(selectedSheetIndexes);

            // Load the workbook using the specified LoadOptions
            Workbook workbook = new Workbook("Template.xlsx", loadOptions);

            // Demonstrate that only the selected sheets are loaded
            Console.WriteLine("Number of worksheets loaded: " + workbook.Worksheets.Count);
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Console.WriteLine($"Sheet {i}: {workbook.Worksheets[i].Name}");
            }

            // (Optional) Save the workbook if further processing is needed
            // workbook.Save("FilteredWorkbook.xlsx");
        }
    }
}
