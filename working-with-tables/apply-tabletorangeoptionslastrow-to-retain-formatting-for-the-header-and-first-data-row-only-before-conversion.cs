// Title: Convert Aspose.Cells Table to a Range while preserving header and first data row (TableToRangeOptions.LastRow)
// Description: Creates a workbook, adds a table, sets TableToRangeOptions.LastRow to 1 (zero‑based) so only the header row and the first data row keep their formatting, converts the table to a normal range, and saves the file.
// Keywords: Aspose.Cells | TableToRangeOptions | LastRow | C# | .NET | convert table to range | preserve header formatting | first data row | ListObject | Excel table conversion
// Common Searches: Aspose.Cells TableToRangeOptions LastRow example | keep header and first row when converting table to range | convert ListObject to range without losing formatting | C# Aspose.Cells table to range conversion
// Developer Intent: Convert an Aspose.Cells ListObject to a regular range while retaining formatting only for the header row and the first data row.
// Use Cases: Flatten a styled table for printing, keeping only the header and a sample row visible. | Export a subset of a table for custom calculations, discarding extra rows but preserving top‑two‑row styling. | Create a template that converts tables to ranges yet leaves the header and first row editable with original styles.
// AI Prompts: Show how to use TableToRangeOptions.LastRow to keep the header and first data row when converting a table to a range in Aspose.Cells for .NET. | Provide a C# code snippet that converts a ListObject to a range and retains formatting only for rows 0 and 1. | Explain the difference between setting TableToRangeOptions.LastRow to 0 versus 1 during table‑to‑range conversion.

using Aspose.Cells;
using Aspose.Cells.Tables;

// Creates a workbook, adds a table, sets TableToRangeOptions.LastRow to 1 (zero‑based) so only the header row and the first data row keep their formatting, converts the table to a normal range, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate header row (row 0)
        cells[0, 0].PutValue("Header1");
        cells[0, 1].PutValue("Header2");

        // Populate first data row (row 1)
        cells[1, 0].PutValue("Data1");
        cells[1, 1].PutValue(100);

        // Populate additional data rows (rows 2‑4)
        for (int r = 2; r < 5; r++)
        {
            cells[r, 0].PutValue($"Data{r}");
            cells[r, 1].PutValue(r * 10);
        }

        // Add a table that spans rows 0‑4 and columns 0‑1
        int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
        ListObject table = sheet.ListObjects[tableIndex];
        table.TableStyleType = TableStyleType.TableStyleMedium2;

        // Configure conversion options to keep only the header row and the first data row
        TableToRangeOptions options = new TableToRangeOptions
        {
            // LastRow is zero‑based; setting it to 1 retains rows 0 and 1
            LastRow = 1
        };

        // Convert the table to a normal range using the specified options
        table.ConvertToRange(options);

        // Save the workbook
        workbook.Save("TableToRangeHeaderAndFirstRow.xlsx");
    }
}
