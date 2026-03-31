using System;
using Aspose.Cells;

namespace AsposeCellsReadOnlyExample
{
    class Program
    {
        static void Main()
        {
            // Path to the large Excel file
            string filePath = "LargeWorkbook.xlsx";

            // Configure load options for minimal memory usage
            LoadOptions loadOptions = new LoadOptions();
            // Prefer memory‑efficient mode (or use FileCache for even lower memory)
            loadOptions.MemorySetting = MemorySetting.MemoryPreference;
            // Do not keep unparsed data (e.g., drawings, comments) in memory
            loadOptions.KeepUnparsedData = false;
            // Skip formula parsing on open – reduces processing overhead
            loadOptions.ParsingFormulaOnOpen = false;

            // Load the workbook in read‑only, memory‑optimized mode
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Example read‑only operation: display the value of cell A1 in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("A1 value: " + sheet.Cells["A1"].StringValue);

            // No modifications are performed; the workbook remains effectively read‑only.
        }
    }
}