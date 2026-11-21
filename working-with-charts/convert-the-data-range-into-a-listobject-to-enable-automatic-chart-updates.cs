// Title: Convert a worksheet range to an Excel Table (ListObject) using Aspose.Cells for .NET to enable automatic chart updates
// Common Searches: Aspose.Cells C# convert A1:C10 range to ListObject for chart refresh | Create Excel table from a cell range with Aspose.Cells .NET | Programmatically add a ListObject with headers using Aspose.Cells | How to apply a built‑in table style to a ListObject in Aspose.Cells C# | Enable dynamic chart data source by converting a range to a table in Aspose.Cells
// Tags: Aspose.Cells create ListObject from range | Aspose.Cells Excel table for dynamic chart | Aspose.Cells apply built‑in table style | Aspose.Cells convert CellArea to table address | Aspose.Cells automatic chart refresh using ListObject

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, defines the A1:C10 range on the first worksheet, adds a ListObject named "MyTable" with headers, applies the TableStyleMedium9 style, and saves the workbook, allowing any linked charts to update automatically when the data changes.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the data range that should become a ListObject (A1:C10)
                CellArea dataRange = new CellArea
                {
                    StartRow = 0,      // Row 1 (zero‑based)
                    StartColumn = 0,   // Column A
                    EndRow = 9,        // Row 10
                    EndColumn = 2      // Column C
                };

                // Convert CellArea to an Excel range string (e.g., "A1:C10")
                string startCell = CellsHelper.CellIndexToName(dataRange.StartRow, dataRange.StartColumn);
                string endCell = CellsHelper.CellIndexToName(dataRange.EndRow, dataRange.EndColumn);
                string rangeAddress = $"{startCell}:{endCell}";

                // Add a ListObject (Excel Table) based on the defined range
                int listObjectIndex = worksheet.ListObjects.Add("MyTable", rangeAddress, true);
                ListObject table = worksheet.ListObjects[listObjectIndex];

                // Apply a built‑in table style
                table.TableStyleType = TableStyleType.TableStyleMedium9;

                // Save the workbook with the new ListObject
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
