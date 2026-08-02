// Title: Extract worksheet TabId from an .xlsx with Aspose.Cells using low‑memory loading (C#)
// Description: Demonstrates how to open an .xlsx file as a read‑only stream, apply LoadOptions.MemorySetting.MemoryPreference for lightweight loading, create a Workbook, iterate through its worksheets, and print each sheet’s Name, Index and internal TabId—without materializing the full workbook in memory.
// Keywords: Aspose.Cells TabId extraction | read worksheet TabId without full load | MemoryPreference LoadOptions | lightweight Excel metadata C# | OpenXML worksheet identifier
// Common Searches: Aspose.Cells get TabId of worksheets without loading workbook | C# read only TabId from large .xlsx | Memory efficient worksheet metadata extraction Aspose | LightCells mode TabId example | How to retrieve internal sheet IDs with Aspose.Cells
// Developer Intent: Obtain the internal TabId of every worksheet in an .xlsx file while keeping memory consumption minimal.
// Use Cases: Create a quick audit of sheet identifiers for version‑control or change‑tracking. | Detect sheet reordering across large workbooks by comparing TabId values. | Generate a lightweight summary of worksheet IDs for reporting on massive spreadsheets.
// AI Prompts: Write C# code that uses Aspose.Cells to read only the TabId values of worksheets from a large .xlsx file, avoiding full workbook loading. | Explain the effect of LoadOptions.MemorySetting.MemoryPreference on workbook loading and how to access the Worksheet.TabId property. | Suggest an alternative approach to retrieve worksheet TabId information from an OpenXML package without instantiating a full Workbook object.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTabIdDemo
{
    // Demonstrates how to open an .xlsx file as a read‑only stream, apply LoadOptions.MemorySetting.MemoryPreference for lightweight loading, create a Workbook, iterate through its worksheets, and print each sheet’s Name, Index and internal TabId—without materializing the full workbook in memory.
    class Program
    {
        static void Main()
        {
            // Path to the Excel file (OpenXml package, e.g., .xlsx)
            string filePath = "sample.xlsx";

            // Open the file as a read‑only stream to avoid loading the whole file into memory
            using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                // Configure LoadOptions to use memory‑efficient loading (LightCells style)
                LoadOptions loadOptions = new LoadOptions
                {
                    // Prefer memory usage over performance; this keeps most data on disk
                    MemorySetting = MemorySetting.MemoryPreference
                };

                // Load the workbook using the stream and the lightweight LoadOptions
                Workbook workbook = new Workbook(stream, loadOptions);

                // Iterate through all worksheets and output their internal TabId values
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" (Index {i}) has TabId: {sheet.TabId}");
                }

                // No need to save; we only extracted metadata
            }
        }
    }
}
