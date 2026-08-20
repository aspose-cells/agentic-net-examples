// Title: C# – Extract worksheet TabId from an XLSX using Aspose.Cells with low‑memory loading
// Description: Demonstrates how to open an XLSX file with LoadOptions.MemorySetting = MemoryPreference, keep only the workbook structure in memory, iterate through each Worksheet, read its TabId property, and output the sheet name, index and TabId. No cell data is materialized and the workbook is not saved.
// Keywords: Aspose.Cells TabId extraction | C# low memory Excel loading | LoadOptions.MemoryPreference | read worksheet metadata OpenXML | XLSX TabId without full workbook | Aspose.Cells memory efficient | Excel worksheet identifier C# | US developers Aspose.Cells
// Common Searches: Aspose.Cells get worksheet TabId without loading whole file | C# low memory read Excel sheet identifiers | How to retrieve TabId from large XLSX using Aspose | MemoryPreference load Excel metadata only | Extract worksheet IDs from OpenXML with Aspose.Cells
// Developer Intent: Obtain the TabId of every worksheet in an XLSX file while avoiding full workbook materialization.
// Use Cases: Validate sheet identifiers in massive workbooks for version control. | Build a name‑to‑TabId map for synchronization with external systems. | Perform quick metadata scans of large Excel files on memory‑constrained servers.
// AI Prompts: Generate C# code that lists each worksheet’s TabId from a large XLSX using Aspose.Cells with MemoryPreference. | Explain the impact of LoadOptions.MemorySetting.MemoryPreference on workbook loading and how to access TabId. | Adapt the sample to export worksheet names and TabIds to a CSV file.

using System;
using Aspose.Cells;
using Aspose.Cells.Metadata;

namespace AsposeCellsTabIdExtractor
{
    // Demonstrates how to open an XLSX file with LoadOptions.MemorySetting = MemoryPreference, keep only the workbook structure in memory, iterate through each Worksheet, read its TabId property, and output the sheet name, index and TabId. No cell data is materialized and the workbook is not saved.
    class Program
    {
        static void Main()
        {
            // Path to the Excel file (OpenXml package)
            string filePath = "sample.xlsx";

            // ------------------------------------------------------------
            // Load the workbook with minimal memory usage.
            // ------------------------------------------------------------
            // LoadOptions with MemoryPreference tells Aspose.Cells to avoid loading
            // the entire workbook into memory. This is suitable for extracting
            // lightweight metadata such as the worksheet TabId.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                // Prefer memory-efficient loading; the workbook structure is still
                // available, but cell data is not fully materialized.
                MemorySetting = MemorySetting.MemoryPreference
            };

            // Create the workbook instance using the constructor that accepts a file path and LoadOptions.
            Workbook workbook = new Workbook(filePath, loadOptions);

            // ------------------------------------------------------------
            // Iterate through worksheets and read the TabId property.
            // ------------------------------------------------------------
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                int tabId = sheet.TabId; // Internal identifier for the sheet
                Console.WriteLine($"Worksheet \"{sheet.Name}\" (Index {i}) has TabId: {tabId}");
            }

            // ------------------------------------------------------------
            // No need to save the workbook because we only read metadata.
            // ------------------------------------------------------------
        }
    }
}
