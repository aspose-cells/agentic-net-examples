// Title: C# AspNet Cells – Convert ListObject to Range, Keep First Three Rows Styling, Export as ODS
// Description: Shows how to build a workbook, apply a light‑gray bold style to the header and two data rows, turn a ListObject into a plain range without losing that style, and write the file as an ODS document with OdsSaveOptions.
// Keywords: Aspose.Cells C# | ListObject to range | preserve row style | ODS export | OdsSaveOptions example | convert table to range | apply row formatting | LibreOffice compatibility | OpenOffice ODS | Aspose.Cells sample code
// Common Searches: Aspose.Cells convert table to range C# | keep row formatting after ListObject conversion | save workbook as ODS using Aspose.Cells | apply style to first rows Aspose.Cells | how to export styled data to ODS
// Developer Intent: Flatten a ListObject while retaining the formatting of selected rows and generate an ODS file.
// Use Cases: Create a report where the header and top rows stay highlighted after removing the table structure for broader tool support. | Export a styled spreadsheet to ODS for seamless opening in LibreOffice or OpenOffice. | Programmatically prepare data, apply custom row styles, convert the table to a range for downstream processing, and save it in ODS format.
// AI Prompts: Generate C# code with Aspose.Cells that converts a ListObject to a normal range, retains the first three rows' styling, and saves the workbook as ODS. | Explain the effect of StyleFlag.All when applying a Style to rows in Aspose.Cells. | Suggest ways to preserve column widths and other layout attributes when turning a table into a range before ODS export.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Ods;

namespace AsposeCellsTableToRangeOds
{
    // Shows how to build a workbook, apply a light‑gray bold style to the header and two data rows, turn a ListObject into a plain range without losing that style, and write the file as an ODS document with OdsSaveOptions.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (5 rows, 3 columns)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells[i - 1, 0].PutValue(i - 1);                     // ID
                sheet.Cells[i - 1, 1].PutValue($"Person {i - 1}");        // Name
                sheet.Cells[i - 1, 2].PutValue(50 + i * 10);              // Score
            }

            // Create a ListObject (table) that covers the data range A1:C5
            int tableIndex = sheet.ListObjects.Add("A1", "C5", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Apply custom formatting to the first three rows (header + two data rows)
            // Create a style with a light gray background and bold font
            Style rowStyle = workbook.CreateStyle();
            rowStyle.ForegroundColor = System.Drawing.Color.LightGray;
            rowStyle.Pattern = BackgroundType.Solid;
            rowStyle.Font.IsBold = true;

            // Define a StyleFlag to apply all formatting attributes
            StyleFlag flag = new StyleFlag
            {
                All = true
            };

            // Apply the style to rows 0, 1, and 2 (zero‑based indices)
            sheet.Cells.ApplyRowStyle(0, rowStyle, flag);
            sheet.Cells.ApplyRowStyle(1, rowStyle, flag);
            sheet.Cells.ApplyRowStyle(2, rowStyle, flag);

            // Convert the table to a normal range while keeping the existing formatting
            table.ConvertToRange();

            // Save the workbook as ODS using OdsSaveOptions
            OdsSaveOptions odsOptions = new OdsSaveOptions();
            // (Optional) Set any additional ODS options here, e.g., odsOptions.IgnorePivotTables = true;

            workbook.Save("TableConvertedToRange.ods", odsOptions);
        }
    }
}
