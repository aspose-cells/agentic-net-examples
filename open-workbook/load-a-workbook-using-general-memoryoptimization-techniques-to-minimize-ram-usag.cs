using System;
using Aspose.Cells;

namespace AsposeCellsMemoryOptimizedLoad
{
    public class Program
    {
        public static void Main()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Configure load options for memory optimization
            LoadOptions loadOptions = new LoadOptions();
            // Prefer memory‑efficient mode (compact data representation)
            loadOptions.MemorySetting = MemorySetting.MemoryPreference;
            // Unparsed data (styles, comments, etc.) is not needed for read‑only scenarios
            loadOptions.KeepUnparsedData = false;

            // Load the workbook using the optimized options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Example read operation: display the value of cell A1 from the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("A1 value: " + sheet.Cells["A1"].StringValue);

            // Save the workbook (optional, demonstrates that the workbook remains functional)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook loaded and saved with memory‑optimized settings.");
        }
    }
}