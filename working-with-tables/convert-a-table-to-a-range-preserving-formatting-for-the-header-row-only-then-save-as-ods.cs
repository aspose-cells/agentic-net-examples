// Title: Aspose.Cells C# – Convert Table to Range, Preserve Header Formatting, Save as ODS
// Description: Creates a workbook, styles the header row (bold, light‑gray), fills data rows, defines a ListObject table, saves the header style, converts the table to a plain range, reapplies the header style, and exports the sheet to ODS using the LibreOffice generator.
// Keywords: Aspose.Cells C# convert table to range | preserve header style Aspose.Cells | ListObject ConvertToRange example | ODS export Aspose.Cells | LibreOffice OdsSaveOptions | C# Excel table to range | Aspose.Cells header formatting
// Common Searches: convert ListObject to range Aspose.Cells C# | keep header formatting after ConvertToRange | save Aspose.Cells workbook as ODS | LibreOffice generator ODS Aspose.Cells | reapply cell style after table conversion
// Developer Intent: Turn an Aspose.Cells table into a normal range while retaining the header row’s visual style and export the result as an ODS file.
// Use Cases: Modify a styled Excel table as a plain range without losing header appearance. | Generate ODS files compatible with LibreOffice after table conversion. | Reapply a saved Style to the header row to ensure consistent formatting post‑conversion.
// AI Prompts: Generate C# code with Aspose.Cells that converts a ListObject to a range and keeps only the header row’s style. | Show how to export an Aspose.Cells workbook to ODS using the LibreOffice generator after converting a table to a range. | Provide an example of saving a header Style before ConvertToRange and reapplying it to the first row in C#.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Ods;

// Creates a workbook, styles the header row (bold, light‑gray), fills data rows, defines a ListObject table, saves the header style, converts the table to a plain range, reapplies the header style, and exports the sheet to ODS using the LibreOffice generator.
class TableToRangePreserveHeader
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add header values
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["C1"].PutValue("Score");
        sheet.Cells["D1"].PutValue("Date");

        // Define header formatting (bold font, light gray background)
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.ForegroundColor = Color.LightGray;
        headerStyle.Pattern = BackgroundType.Solid;

        // Apply the header style to the header row (A1:D1)
        sheet.Cells.CreateRange("A1:D1").ApplyStyle(headerStyle, new StyleFlag { All = true });

        // Populate some data rows (rows 2 to 5)
        for (int row = 2; row <= 5; row++)
        {
            sheet.Cells[row - 1, 0].PutValue(row - 1);                     // ID
            sheet.Cells[row - 1, 1].PutValue($"Person {row - 1}");        // Name
            sheet.Cells[row - 1, 2].PutValue((row - 1) * 10);             // Score
            sheet.Cells[row - 1, 3].PutValue(DateTime.Today.AddDays(row - 2)); // Date
        }

        // Create a table (ListObject) that includes the header and data rows
        int tableIdx = sheet.ListObjects.Add("A1", "D5", true);
        ListObject table = sheet.ListObjects[tableIdx];
        table.TableStyleType = TableStyleType.TableStyleMedium2; // optional visual style

        // Save the current header style so it can be reapplied after conversion
        Style savedHeaderStyle = sheet.Cells["A1"].GetStyle();

        // Convert the table to a normal range
        table.ConvertToRange();

        // Reapply the saved header style to the first row (header) after conversion
        for (int col = 0; col < 4; col++)
        {
            sheet.Cells[0, col].SetStyle(savedHeaderStyle);
        }

        // Prepare ODS save options (using LibreOffice generator)
        OdsSaveOptions odsOptions = new OdsSaveOptions();
        odsOptions.GeneratorType = OdsGeneratorType.LibreOffice;

        // Save the workbook as an ODS file
        workbook.Save("TableConverted.ods", odsOptions);
    }
}
