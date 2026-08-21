// Title: C# – Convert Aspose.Cells ListObject (Table) to a Range and Copy to Another Worksheet with Range.Copy
// Description: This example creates a workbook, adds sample data, defines a ListObject covering A1:B4, converts the table to a regular range using ListObject.ConvertToRange, creates a matching source range, defines a destination range starting at D1 on a new sheet, copies the source range to the destination with Range.Copy, and saves the file as TableConvertedAndCopied.xlsx.
// Keywords: Aspose.Cells | C# | ListObject | ConvertToRange | Range.Copy | copy range between worksheets | convert table to range | Aspose.Cells example | copy data to another sheet | Aspose.Cells API
// Common Searches: How to convert an Aspose.Cells table to a range in C# | Aspose.Cells Range.Copy between worksheets | Copy ListObject data to another sheet using Aspose.Cells | Convert ListObject to range and duplicate in C# | Aspose.Cells copy range to specific cell D1
// Developer Intent: Convert a ListObject to a normal range and duplicate that range on a different worksheet.
// Use Cases: Create a summary sheet by copying raw table data without table formatting | Transfer data to a report layout where the table must start at a specific column | Programmatically move converted ranges for further calculations or charting
// AI Prompts: Provide C# code to convert an Aspose.Cells ListObject to a range and copy it to another worksheet. | Explain how ListObject.ConvertToRange and Range.Copy work together in Aspose.Cells. | Generate error‑handling for copying a converted table range between sheets in Aspose.Cells C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsTableToRangeCopy
{
    // This example creates a workbook, adds sample data, defines a ListObject covering A1:B4, converts the table to a regular range using ListObject.ConvertToRange, creates a matching source range, defines a destination range starting at D1 on a new sheet, copies the source range to the destination with Range.Copy, and saves the file as TableConvertedAndCopied.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // ---------- Source worksheet ----------
                Worksheet srcSheet = workbook.Worksheets[0];
                srcSheet.Name = "Source";

                // Fill sample data for the table (A1:B4)
                srcSheet.Cells["A1"].PutValue("ID");
                srcSheet.Cells["B1"].PutValue("Name");
                srcSheet.Cells["A2"].PutValue(1);
                srcSheet.Cells["B2"].PutValue("John");
                srcSheet.Cells["A3"].PutValue(2);
                srcSheet.Cells["B3"].PutValue("Mary");
                srcSheet.Cells["A4"].PutValue(3);
                srcSheet.Cells["B4"].PutValue("Bob");

                // Add a ListObject (table) covering the data range
                int tableIndex = srcSheet.ListObjects.Add("A1", "B4", true);
                ListObject table = srcSheet.ListObjects[tableIndex];

                // Convert the table to a normal range
                table.ConvertToRange();

                // Define the source range that was previously the table
                // (first row = 0, first column = 0, rows = 4, columns = 2)
                AsposeRange sourceRange = srcSheet.Cells.CreateRange(0, 0, 4, 2);

                // ---------- Destination worksheet ----------
                Worksheet destSheet = workbook.Worksheets.Add("Destination");

                // Create a destination range of the same size starting at cell D1 (column index 3)
                AsposeRange destRange = destSheet.Cells.CreateRange(0, 3, 4, 2); // D1:E4

                // Copy the source range to the destination range
                sourceRange.Copy(destRange);

                // Save the workbook
                workbook.Save("TableConvertedAndCopied.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
