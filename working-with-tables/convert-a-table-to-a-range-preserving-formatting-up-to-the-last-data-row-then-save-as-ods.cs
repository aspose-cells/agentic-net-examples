// Title: Convert Aspose.Cells ListObject to a Range with Formatting and Export to ODS (C#)
// Description: Demonstrates how to create a workbook, style a header, fill data rows, define a ListObject, detect the last populated row in column A, convert the table to a normal range while preserving its formatting up to that row using TableToRangeOptions, and finally save the result as an ODS file with OdsSaveOptions.
// Keywords: Aspose.Cells C# table to range | ListObject conversion | preserve formatting Aspose.Cells | TableToRangeOptions LastRow | save as ODS Aspose.Cells | OpenDocument Spreadsheet export | Excel table to range example
// Common Searches: Aspose.Cells convert ListObject to range C# | keep table formatting when converting to range Aspose | how to export workbook as ODS using Aspose.Cells | determine last data row in Aspose.Cells column | TableToRangeOptions example
// Developer Intent: Transform an Excel ListObject into a regular range, retain its styling up to the final data row, and export the workbook as an ODS document.
// Use Cases: Prepare a spreadsheet for ODS export when the source uses Excel tables that must become plain ranges. | Apply operations that are unsupported on tables (e.g., custom merged cells) after converting the table while keeping header styles. | Generate reports that require precise row limits, avoiding empty rows by using the last populated row as a conversion boundary.
// AI Prompts: Write C# code with Aspose.Cells to convert a ListObject to a range, preserve formatting up to the last data row, and save the file as ODS. | Explain the role of TableToRangeOptions.LastRow and show how to obtain the last non‑empty row in a specific column using Aspose.Cells. | Provide an alternative snippet that performs the same conversion but saves the workbook as XLSX instead of ODS.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Ods;   // Namespace for OdsSaveOptions
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

namespace AsposeCellsTableToRangeOds
{
    // Demonstrates how to create a workbook, style a header, fill data rows, define a ListObject, detect the last populated row in column A, convert the table to a normal range while preserving its formatting up to that row using TableToRangeOptions, and finally save the result as an ODS file with OdsSaveOptions.
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

                // Populate sample data (including some formatting)
                // Header row
                cells["A1"].PutValue("ID");
                cells["B1"].PutValue("Name");
                cells["C1"].PutValue("Score");

                // Apply a simple style to the header
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.ForegroundColor = Color.LightGray;
                headerStyle.Pattern = BackgroundType.Solid;

                // Apply the style to the header range A1:C1
                AsposeRange headerRange = cells.CreateRange("A1:C1");
                StyleFlag flag = new StyleFlag { All = true };
                headerRange.ApplyStyle(headerStyle, flag);

                // Data rows
                for (int i = 2; i <= 10; i++)
                {
                    cells[i - 1, 0].PutValue(i - 1);                 // ID
                    cells[i - 1, 1].PutValue($"Person {i - 1}");    // Name
                    cells[i - 1, 2].PutValue(50 + i);               // Score
                }

                // Create a ListObject (table) that covers the data including the header
                int tableIndex = sheet.ListObjects.Add("A1", "C10", true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Determine the last row that contains data in the first column (ID column)
                int lastDataRow = cells.GetLastDataRow(0); // zero‑based index

                // Convert the table to a normal range, preserving formatting up to the last data row
                TableToRangeOptions options = new TableToRangeOptions
                {
                    LastRow = lastDataRow   // Convert only rows 0..lastDataRow
                };
                table.ConvertToRange(options);

                // Save the workbook as ODS using OdsSaveOptions
                OdsSaveOptions odsOptions = new OdsSaveOptions();
                workbook.Save("TableConvertedToRange.ods", odsOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
