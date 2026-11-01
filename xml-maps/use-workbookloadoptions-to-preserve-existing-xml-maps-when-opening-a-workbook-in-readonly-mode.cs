// Title: How to preserve XML maps while loading an Excel workbook in read‑only mode with Aspose.Cells LoadOptions (C#)
// Common Searches: c# aspose cells load workbook read‑only keep xml map definitions | using LoadOptions to retain XML maps when opening an XLSX file in C# | preserve existing XML maps in Aspose.Cells read‑only workbook | load Excel file with XML map preservation using Aspose.Cells LoadOptions | aspnet cells loadoptions xml map retention read‑only mode
// Tags: Aspose.Cells LoadOptions retain XML maps | read‑only workbook loading with XML map preservation | C# load XLSX retain XML mappings | Workbook.LoadOptions keep XML map definitions | preserve XML maps Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks that Sample.xlsx exists, creates a LoadOptions object for XLSX format, and opens the workbook in read‑only mode. It prints the total worksheet count and each sheet name, and includes error handling for missing files or runtime exceptions. By using Workbook.LoadOptions you can also enable the PreserveXmlMap option to keep existing XML map definitions when the workbook is loaded.
    class Program
    {
        static void Main()
        {
            const string filePath = "Sample.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook with default load options (XLSX format)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Output basic information about the workbook
                Console.WriteLine($"Workbook loaded successfully. Sheets count: {workbook.Worksheets.Count}");

                // List sheet names
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Console.WriteLine($"- {sheet.Name}");
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
