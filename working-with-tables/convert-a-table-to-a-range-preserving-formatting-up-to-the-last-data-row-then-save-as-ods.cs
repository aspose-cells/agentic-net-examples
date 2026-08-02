// Title: C# AspNet Cells – Convert ListObject to Range (keep formatting) and Export to ODS
// Description: Shows how to build a workbook, add a ListObject (table), locate the last populated row in column A, transform the table into a normal range while retaining its style via TableToRangeOptions, and finally write the file as an OpenDocument Spreadsheet (ODS) using OdsSaveOptions.
// Keywords: Aspose.Cells C# convert table to range | TableToRangeOptions LastRow | preserve table styling Aspose.Cells | save workbook as ODS | OdsSaveOptions C# | ListObject to range conversion | export to LibreOffice ODS | detect last data row Aspose.Cells
// Common Searches: Aspose.Cells convert ListObject to range C# | How to keep table formatting when converting to range Aspose | Save Aspose.Cells workbook as ODS file | Get last data row index column A Aspose.Cells | TableToRangeOptions example C#
// Developer Intent: I need to change a ListObject into a regular cell range, retain its visual formatting up to the final data row, and then export the workbook as an ODS file.
// Use Cases: Prepare data for calculations that require plain ranges after styling a table | Generate ODS reports for LibreOffice while preserving the original table appearance | Remove table functionality before printing or sharing without losing formatting | Integrate with workflows that only accept ODS files but start with Excel tables
// AI Prompts: Write C# code using Aspose.Cells to convert a ListObject to a regular range with TableToRangeOptions, preserving formatting up to the last data row. | Show how to detect the last populated row in a specific column and apply it to TableToRangeOptions. | Provide an example of saving an Aspose.Cells workbook as ODS using OdsSaveOptions, disabling pivot table export. | Explain why converting a table to a range might be necessary before performing certain cell‑level operations.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Shows how to build a workbook, add a ListObject (table), locate the last populated row in column A, transform the table into a normal range while retaining its style via TableToRangeOptions, and finally write the file as an OpenDocument Spreadsheet (ODS) using OdsSaveOptions.
    public class TableToRangeAndSaveOds
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including header row)
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Score");

            // Data rows
            for (int row = 2; row <= 10; row++)
            {
                cells[row - 1, 0].PutValue(row - 1);                     // ID
                cells[row - 1, 1].PutValue($"Person {row - 1}");        // Name
                cells[row - 1, 2].PutValue((row - 1) * 10);             // Score
            }

            // Create a ListObject (table) that covers the whole data range (A1:C10)
            int tableIndex = sheet.ListObjects.Add("A1", "C10", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Optional: set a style for visual verification
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Determine the last row that contains data in the first column (ID column)
            int lastDataRow = sheet.Cells.GetLastDataRow(0);

            // Convert the table to a normal range, preserving formatting up to the last data row
            TableToRangeOptions options = new TableToRangeOptions
            {
                // LastRow expects a zero‑based index.
                LastRow = lastDataRow
            };
            table.ConvertToRange(options);

            // Save the workbook as ODS using OdsSaveOptions
            OdsSaveOptions odsOptions = new OdsSaveOptions
            {
                // Example: ignore pivot tables (not needed here but shows option usage)
                IgnorePivotTables = true
            };

            workbook.Save("TableConvertedToRange.ods", odsOptions);
        }
    }
}
