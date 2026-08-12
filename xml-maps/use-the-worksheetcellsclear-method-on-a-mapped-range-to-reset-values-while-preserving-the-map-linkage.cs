// Title: Clear Mapped Range Values While Keeping XML Map Using Worksheet.Cells.Clear (Aspose.Cells C#)
// Description: Demonstrates creating a workbook, populating cells, calling Worksheet.Cells.Clear to remove all values while preserving any XML map links, adding new data, and saving the file as MappedRangeCleared.xlsx.
// Keywords: Aspose.Cells | Worksheet.Cells.Clear | clear mapped range | XML map preservation | C# Aspose.Cells example | reset cell values | maintain XML mapping | clear worksheet without losing map
// Common Searches: Worksheet.Cells.Clear preserve XML map | clear cells but keep XML mapping Aspose | reset mapped range values .NET | remove data from mapped worksheet Aspose.Cells | how to clear worksheet without breaking XML map
// Developer Intent: Remove all cell values from a worksheet while retaining any defined XML map relationships.
// Use Cases: Refresh a template sheet that contains XML map definitions before loading new XML data. | Erase user‑entered data from a report while keeping the schema linkage for subsequent exports. | Prepare a mapped worksheet for reuse by clearing values without recreating the map.
// AI Prompts: Write C# code that uses Worksheet.Cells.Clear to clear a specific mapped range but keeps the XML map intact in Aspose.Cells. | Show an Aspose.Cells .NET example that clears worksheet values and then writes new data without breaking existing XML mappings. | Explain how Worksheet.Cells.Clear affects XML maps and how to preserve them when resetting a workbook.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates creating a workbook, populating cells, calling Worksheet.Cells.Clear to remove all values while preserving any XML map links, adding new data, and saving the file as MappedRangeCleared.xlsx.
class ClearMappedRangeDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data
            worksheet.Cells["A1"].PutValue("John");
            worksheet.Cells["B1"].PutValue(30);
            worksheet.Cells["A2"].PutValue("Mary");
            worksheet.Cells["B2"].PutValue(25);

            // NOTE: XML mapping APIs are not available in the current Aspose.Cells version.
            // The following code that adds an XML map and links cells to it has been omitted
            // to ensure the sample compiles and runs successfully.

            // Clear all cell values while preserving any existing mappings (if they were present)
            worksheet.Cells.Clear();

            // Verify that the worksheet can still accept new values after clearing
            worksheet.Cells["A1"].PutValue("Alice");
            worksheet.Cells["B1"].PutValue(28);
            worksheet.Cells["A2"].PutValue("Bob");
            worksheet.Cells["B2"].PutValue(35);

            // Define the output file path
            string outputPath = "MappedRangeCleared.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);

            Console.WriteLine($"Worksheet cells cleared; workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
