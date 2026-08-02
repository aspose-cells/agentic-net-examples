// Title: C# – Convert Every Table to a Range with TableToRangeOptions Using Aspose.Cells
// Description: Loads a workbook, walks through each worksheet, retrieves all ListObjects, and converts every table to a normal cell range with a custom TableToRangeOptions (e.g., setting LastRow). The conversion runs backward to keep the collection stable, then the workbook is saved.
// Keywords: Aspose.Cells C# convert table to range | TableToRangeOptions LastRow | ListObjectCollection iteration | convert ListObjects Aspose.Cells | batch table conversion .NET | Aspose.Cells workbook tables | C# Excel table to range
// Common Searches: Aspose.Cells convert all tables to ranges C# | TableToRangeOptions example Aspose.Cells | Iterate worksheets and ListObjects Aspose.Cells | How to remove tables after conversion Aspose.Cells | Backward loop ListObjectCollection ConvertToRange
// Developer Intent: Transform each ListObject in a workbook into a plain range using custom conversion options.
// Use Cases: Strip table formatting while preserving data before exporting to CSV or plain‑text formats. | Limit the converted area to the actual last row of each table to avoid trailing blanks. | Prepare workbooks for older Excel versions that do not support structured tables.
// AI Prompts: Write C# code that loops through all worksheets in a workbook and converts each ListObject to a range with TableToRangeOptions, setting LastRow to the table's EndRow. | Explain why a reverse for‑loop is required when calling ConvertToRange on a ListObjectCollection and show how to also configure FirstColumn in TableToRangeOptions. | Provide a sample that converts tables to ranges, keeps formulas and formatting intact, and then saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableConversion
{
    // Loads a workbook, walks through each worksheet, retrieves all ListObjects, and converts every table to a normal cell range with a custom TableToRangeOptions (e.g., setting LastRow). The conversion runs backward to keep the collection stable, then the workbook is saved.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of tables (ListObjects) on the current worksheet
                ListObjectCollection tables = sheet.ListObjects;

                // Iterate backwards because ConvertToRange removes the table from the collection
                for (int i = tables.Count - 1; i >= 0; i--)
                {
                    ListObject table = tables[i];

                    // Create custom options for conversion
                    TableToRangeOptions options = new TableToRangeOptions
                    {
                        // Example: convert only up to the current last row of the table
                        // (you can set any row index you need)
                        LastRow = table.EndRow
                    };

                    // Convert the table to a normal range using the custom options
                    table.ConvertToRange(options);
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}
