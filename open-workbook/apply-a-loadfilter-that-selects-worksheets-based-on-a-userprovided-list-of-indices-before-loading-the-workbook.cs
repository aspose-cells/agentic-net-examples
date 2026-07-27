// Title: C# Example: Load Specific Worksheets by Index with Aspose.Cells Custom LoadFilter
// Description: Shows how to build a CustomLoadFilter that receives an array of zero‑based worksheet indexes, overrides SheetsInLoadingOrder, and plugs into LoadOptions so that Workbook loads only the chosen sheets. The sample prints the loaded sheet count, iterates over each sheet, and optionally saves the filtered workbook, enabling memory‑efficient processing of large Excel files.
// Keywords: Aspose.Cells C# LoadFilter | load worksheets by index | select specific sheets .NET | LoadOptions custom filter | memory efficient Excel loading | filter workbook worksheets | zero based sheet indexes | Aspose.Cells example GitHub | C# Excel sheet selection | Aspose.Cells LoadDataFilterOptions
// Common Searches: Aspose.Cells load only certain worksheets | C# custom LoadFilter sheet index | How to load selected sheets with Aspose.Cells | LoadOptions LoadFilter example .NET | Filter workbook sheets by index Aspose
// Developer Intent: Load a workbook while including only the worksheets whose indexes are supplied by the user.
// Use Cases: Reduce memory consumption by loading just the first and third sheets of a large template. | Reorder sheets during import based on a user‑defined sequence of indexes. | Create a lightweight copy of an Excel file that contains only the worksheets requested by a client.
// AI Prompts: Write a C# snippet that uses Aspose.Cells LoadFilter to load worksheets whose indexes are provided in a list and saves the result to a new file. | Explain how to adapt the CustomLoadFilter to exclude sheets instead of including them, with safe handling of out‑of‑range indexes. | Provide best‑practice code for validating user‑supplied sheet indexes before constructing the LoadFilter to prevent runtime errors.

using System;
using Aspose.Cells;

namespace AsposeCellsLoadFilterExample
{
    // Custom LoadFilter that returns a specific sheet loading order
    // Shows how to build a CustomLoadFilter that receives an array of zero‑based worksheet indexes, overrides SheetsInLoadingOrder, and plugs into LoadOptions so that Workbook loads only the chosen sheets. The sample prints the loaded sheet count, iterates over each sheet, and optionally saves the filtered workbook, enabling memory‑efficient processing of large Excel files.
    public class CustomLoadFilter : LoadFilter
    {
        private readonly int[] _sheetsOrder;

        // Constructor accepts the desired sheet indices
        public CustomLoadFilter(int[] sheetsOrder) : base(LoadDataFilterOptions.All)
        {
            _sheetsOrder = sheetsOrder;
        }

        // Override the read‑only property to supply the sheet order
        public override int[] SheetsInLoadingOrder => _sheetsOrder;
    }

    class Program
    {
        static void Main()
        {
            // User‑provided list of sheet indexes to load (zero‑based)
            int[] selectedSheetIndexes = new int[] { 0, 2 };

            // Create the custom filter with the desired indexes
            CustomLoadFilter loadFilter = new CustomLoadFilter(selectedSheetIndexes);

            // Configure LoadOptions to use the custom filter
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFilter = loadFilter
            };

            // Load the workbook applying the filter
            Workbook workbook = new Workbook("Template.xlsx", loadOptions);

            // Output information about the loaded sheets
            Console.WriteLine("Number of sheets loaded: " + workbook.Worksheets.Count);
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Sheet Index: {sheet.Index}, Name: {sheet.Name}");
            }

            // Save the filtered workbook (optional)
            workbook.Save("FilteredOutput.xlsx");
        }
    }
}
