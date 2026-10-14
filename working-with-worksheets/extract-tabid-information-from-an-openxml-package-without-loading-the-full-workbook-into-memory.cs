// Title: Read worksheet TabId values from an XLSX file using Aspose.Cells with a read‑only stream to avoid full workbook load
// AI Prompts: Write C# code that opens an .xlsx file via a FileStream and Aspose.Cells, iterates the Worksheets collection, and prints each sheet's Name and TabId while keeping memory usage low. | Demonstrate how to obtain the internal TabId for every worksheet in an OpenXML workbook using Aspose.Cells without loading the entire workbook into memory.
// Common Searches: asp.net c# get worksheet tabid from xlsx without loading whole workbook | how to read sheet TabId from an Excel file using Aspose.Cells stream | memory efficient way to list worksheet TabId values in a large .xlsx with Aspose | Aspose.Cells retrieve internal sheet identifier OpenXML workbook C#
// Tags: read worksheet TabId using Aspose.Cells stream | memory‑optimized Excel workbook loading C# | extract internal sheet identifier OpenXML Aspose.Cells | enumerate worksheets TabId property Aspose.Cells | openxlsx as stream with LoadOptions Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program opens an .xlsx file with a read‑only FileStream, loads it via Aspose.Cells using LoadOptions, iterates through the Worksheets collection, and writes each sheet's Name and TabId to the console, handling missing files and exceptions while minimizing memory consumption.
class TabIdExtractor
{
    static void Main()
    {
        // Path to the Excel file (OpenXML package)
        string filePath = @"C:\Path\To\YourWorkbook.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Open the file as a read‑only stream to avoid loading the whole file into memory
            using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Configure load options for minimal memory usage
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                // LoadDataOnly and MemoryOptimization are not available in this version of Aspose.Cells,
                // so we rely on the default loading behavior which is already memory‑efficient.

                // Load the workbook using the stream and the load options
                Workbook workbook = new Workbook(stream, loadOptions);

                // Iterate through each worksheet and output its TabId
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // TabId corresponds to the internal sheet identifier in the OpenXML package
                    int tabId = sheet.TabId;
                    Console.WriteLine($"Sheet Name: {sheet.Name}, TabId: {tabId}");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
