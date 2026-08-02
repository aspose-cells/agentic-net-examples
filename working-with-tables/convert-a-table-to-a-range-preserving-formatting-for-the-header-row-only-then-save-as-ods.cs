// Title: Convert ListObject to Range, Preserve Header Style, Save as ODS (C#)
// Description: Demonstrates how to create a workbook with a styled table, capture the header row formatting, convert the Aspose.Cells ListObject to a plain range, reapply the saved styles, configure OdsSaveOptions (LibreOffice generator), and export the result as an ODS file using .NET.
// Keywords: Aspose.Cells C# convert table to range | preserve header formatting Aspose.Cells | save workbook as ODS | ListObject to range Aspose.Cells | OdsSaveOptions LibreOffice | C# Excel to ODS conversion | Aspose.Cells table styling
// Common Searches: Aspose.Cells convert ListObject to range C# | keep header style after table conversion Aspose.Cells | export Excel workbook to ODS with Aspose.Cells | how to use OdsSaveOptions in .NET | store and reapply cell styles Aspose.Cells
// Developer Intent: Turn a ListObject into a normal range while retaining the header row’s visual style and save the workbook as an ODS document.
// Use Cases: Need to remove table metadata but keep the original header appearance for downstream processing. | Generating ODS files from Excel workbooks that originally contain tables, ensuring consistent styling. | Applying custom ODS save options such as the LibreOffice generator before exporting.
// AI Prompts: Show C# code that converts an Aspose.Cells ListObject to a range, preserves header formatting, and saves the file as ODS. | Explain how to capture cell styles before converting a table to a range and reapply them afterward with Aspose.Cells. | Provide an example of configuring OdsSaveOptions with the LibreOffice generator in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Ods;

namespace AsposeCellsTableToRangeOds
{
    // Demonstrates how to create a workbook with a styled table, capture the header row formatting, convert the Aspose.Cells ListObject to a plain range, reapply the saved styles, configure OdsSaveOptions (LibreOffice generator), and export the result as an ODS file using .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells[i - 1, 0].PutValue(i - 1);                 // ID
                sheet.Cells[i - 1, 1].PutValue($"Person {i - 1}");   // Name
                sheet.Cells[i - 1, 2].PutValue(50 + i * 10);         // Score
            }

            // Apply formatting to the header row (first row)
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;
            // Apply the style to the header cells
            for (int col = 0; col < 3; col++)
            {
                sheet.Cells[0, col].SetStyle(headerStyle);
            }

            // Create a ListObject (table) that includes the header and data
            int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.ShowTableStyleFirstColumn = false; // optional visual tweaks

            // Preserve the header row style before conversion
            // Store the style of each header cell in an array
            Style[] savedHeaderStyles = new Style[3];
            for (int col = 0; col < 3; col++)
            {
                savedHeaderStyles[col] = sheet.Cells[0, col].GetStyle();
            }

            // Convert the table to a normal range (the table object will be removed)
            table.ConvertToRange();

            // Reapply the saved header styles to the first row of the range
            // After conversion the data remains at the same addresses, so we reuse the same cells
            for (int col = 0; col < 3; col++)
            {
                sheet.Cells[0, col].SetStyle(savedHeaderStyles[col]);
            }

            // Prepare ODS save options (optional: set generator type)
            OdsSaveOptions odsOptions = new OdsSaveOptions();
            odsOptions.GeneratorType = OdsGeneratorType.LibreOffice;

            // Save the workbook as ODS
            workbook.Save("TableConvertedToRange.ods", odsOptions);
        }
    }
}
