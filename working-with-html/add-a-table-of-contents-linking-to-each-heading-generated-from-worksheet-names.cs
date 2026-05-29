using System;
using Aspose.Cells;

namespace AsposeCellsTableOfContents
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the worksheets collection
                WorksheetCollection sheets = workbook.Worksheets;

                // Add sample worksheets with distinct names
                sheets.Add("Sales");
                sheets.Add("Inventory");
                sheets.Add("Employees");
                sheets.Add("Summary");

                // Insert a Table of Contents sheet at the first position
                Worksheet tocSheet = sheets.Insert(0, SheetType.Worksheet, "Table of Contents");

                // Populate the TOC with hyperlinks to each worksheet
                int tocRow = 0;
                foreach (Worksheet ws in sheets)
                {
                    // Skip the TOC sheet itself
                    if (ws.Name == "Table of Contents")
                        continue;

                    // Write the worksheet name in column A
                    tocSheet.Cells[tocRow, 0].PutValue(ws.Name);

                    // Add a hyperlink that points to cell A1 of the target worksheet
                    // Hyperlink address format: 'SheetName'!A1
                    string address = $"'{ws.Name}'!A1";

                    // totalRows = 1, totalColumns = 1 for a single-cell hyperlink
                    tocSheet.Hyperlinks.Add(tocRow, 0, 1, 1, address);

                    tocRow++;
                }

                // Save the workbook with the generated Table of Contents
                workbook.Save("TableOfContents.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}