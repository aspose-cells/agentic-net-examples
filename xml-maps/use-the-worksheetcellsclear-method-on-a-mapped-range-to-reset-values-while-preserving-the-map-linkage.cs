// Title: C# – Clear a mapped range with Worksheet.Cells.ClearContents while keeping XML map (Aspose.Cells)
// Description: A .NET example that creates a workbook, writes sample data to A1:B2, defines the same area with a CellArea object, and calls Worksheet.Cells.ClearContents to erase the values while preserving formatting and the XML map binding, then saves the workbook.
// Keywords: Aspose.Cells | Worksheet.Cells.ClearContents | XML map | mapped range | C# example | CellArea | clear cell values | preserve map binding | reset mapped cells | Aspose.Cells .NET
// Common Searches: How to clear values in an XML‑mapped range using Aspose.Cells C# | Worksheet.Cells.ClearContents on a CellArea without breaking XML map | Preserve XML map after clearing cells in Aspose.Cells | Aspose.Cells clear mapped range example | Reset data in mapped cells while keeping map linkage
// Developer Intent: Remove data from a specific XML‑mapped range without disrupting its map association.
// Use Cases: Clear old data from a mapped region before importing new XML into the workbook. | Reset user‑entered values in a template while retaining underlying map bindings. | Programmatically wipe a table area while keeping cell styles and XML map connections intact.
// AI Prompts: Show how to use Worksheet.Cells.ClearContents with a CellArea to clear a mapped range and keep the XML map intact in Aspose.Cells for .NET. | Generate C# code that clears only the values of a mapped range (A1:B2) while preserving formatting and map linkage using Aspose.Cells. | Explain the difference between Clear, ClearContents, and ClearFormats when working with XML‑mapped cells in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsClearMappedRange
{
    // A .NET example that creates a workbook, writes sample data to A1:B2, defines the same area with a CellArea object, and calls Worksheet.Cells.ClearContents to erase the values while preserving formatting and the XML map binding, then saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample data to be cleared later
                cells["A1"].PutValue("John");
                cells["B1"].PutValue(30);
                cells["A2"].PutValue("Mary");
                cells["B2"].PutValue(25);

                // Define the range to clear (A1:B2) using CellArea.
                CellArea clearArea = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 1,
                    EndColumn = 1
                };

                // Clear only the contents; formatting remains intact.
                cells.ClearContents(clearArea);

                // Save the workbook to verify the result.
                workbook.Save("ClearedMappedRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
