// Title: C# – Convert an Excel Table to a Range (first 5 rows) and Export as ODS with Aspose.Cells
// Description: An Aspose.Cells for .NET sample that builds a workbook, fills it with headers and data, applies a light‑yellow style to rows 0‑4, creates a ListObject table, converts the table to a plain range while preserving the first five rows' formatting, and saves the result as an ODS file using OdsSaveOptions.
// Keywords: Aspose.Cells | C# | Convert Table to Range | TableToRangeOptions | preserve row formatting | ODS export | ListObject | Excel to ODS | style first rows | .NET
// Common Searches: Aspose.Cells convert table to range C# example | keep formatting when converting Excel table to range | save workbook as ODS using Aspose.Cells .NET | TableToRangeOptions LastRow usage | export styled Excel rows to ODS format
// Developer Intent: Programmatically transform a ListObject into a regular cell range, retain the formatting of the top five rows, and write the workbook to OpenDocument Spreadsheet (ODS).
// Use Cases: Generate a LibreOffice‑compatible report where only the header and the first few highlighted rows need to remain styled after removing the table structure. | Create a data‑processing template that strips table metadata while preserving background colors for the initial rows before sharing the file with non‑Excel users. | Automate conversion of a styled Excel table to a plain range for downstream analytics, then export the result as ODS for cross‑platform accessibility.
// AI Prompts: Write C# code with Aspose.Cells that converts a ListObject to a range, keeps formatting for rows 0‑4, and saves the workbook as an ODS file. | Explain how TableToRangeOptions.LastRow influences the conversion of an Excel table to a range and how cell styles are retained in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Ods; // optional, OdsSaveOptions is also in Aspose.Cells

// An Aspose.Cells for .NET sample that builds a workbook, fills it with headers and data, applies a light‑yellow style to rows 0‑4, creates a ListObject table, converts the table to a plain range while preserving the first five rows' formatting, and saves the result as an ODS file using OdsSaveOptions.
class ConvertTableToRange
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Fill sample data (10 data rows + header, 3 columns)
        for (int col = 0; col < 3; col++)
        {
            cells[0, col].PutValue($"Header {col + 1}");
        }

        for (int row = 1; row <= 10; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                cells[row, col].PutValue($"R{row}C{col + 1}");
            }
        }

        // Apply a background style to the first five rows (rows 0‑4)
        Style style = workbook.CreateStyle();
        style.ForegroundColor = Color.LightYellow;
        style.Pattern = BackgroundType.Solid;
        StyleFlag flag = new StyleFlag();
        flag.CellShading = true;

        for (int r = 0; r < 5; r++)
        {
            for (int c = 0; c < 3; c++)
            {
                cells[r, c].SetStyle(style, flag);
            }
        }

        // Create a table that spans all data rows (0‑10) and columns (0‑2)
        int tableIdx = sheet.ListObjects.Add(0, 0, 10, 2, true);
        ListObject table = sheet.ListObjects[tableIdx];
        table.TableStyleType = TableStyleType.TableStyleMedium2;

        // Convert the table to a range, keeping only the first five rows
        TableToRangeOptions options = new TableToRangeOptions
        {
            LastRow = 4 // zero‑based index; rows 0‑4 correspond to the first five rows
        };
        table.ConvertToRange(options);

        // Save the workbook as ODS
        OdsSaveOptions odsOptions = new OdsSaveOptions();
        workbook.Save("TableConverted.ods", odsOptions);
    }
}
