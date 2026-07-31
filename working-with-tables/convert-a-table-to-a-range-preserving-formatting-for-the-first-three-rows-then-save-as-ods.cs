// Title: Convert an Excel ListObject to a range, preserve first‑three‑rows formatting, and save as ODS with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, style the first three rows, add a ListObject over A1:C6, convert the table back to a regular range while retaining the applied style, and export the result to an ODS file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells convert ListObject to range | preserve row formatting Aspose.Cells | save workbook as ODS .NET | C# Aspose.Cells table to range example | OdsSaveOptions usage | Excel table to range conversion
// Common Searches: convert Excel table to range with Aspose.Cells and keep formatting | Aspose.Cells .NET export to ODS after removing ListObject | how to keep header style when converting ListObject to range | C# example: ListObject to range then save as ODS | Aspose.Cells preserve cell style during ConvertToRange
// Developer Intent: Transform a ListObject into a normal range without losing the formatting of the first three rows, then generate an ODS file.
// Use Cases: Prepare OpenDocument reports where tables must be flattened but visual styles remain intact. | Share Excel data with partners who require ODS format while retaining custom header and row designs. | Perform downstream data processing that expects plain ranges, not tables, without sacrificing presentation quality.
// AI Prompts: Generate C# code using Aspose.Cells to convert a ListObject to a range, keep the first three rows' formatting, and save the workbook as ODS. | Show how to apply a style to rows 1‑3, create a table over A1:C6, convert the table to a regular range, and export to ODS with Aspose.Cells for .NET. | Explain the effect of ConvertToRange on existing cell styles and how to configure OdsSaveOptions for optimal ODS output.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Ods;

namespace AsposeCellsTableToRangeOds
{
    // Demonstrates how to create a workbook, style the first three rows, add a ListObject over A1:C6, convert the table back to a regular range while retaining the applied style, and export the result to an ODS file using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (header + 5 data rows)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[i - 1, 0].PutValue(i - 1);                     // ID
                sheet.Cells[i - 1, 1].PutValue($"Person {i - 1}");        // Name
                sheet.Cells[i - 1, 2].PutValue(50 + i * 5);               // Score
            }

            // Apply distinct formatting to the first three rows (including header)
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;
            sheet.Cells.CreateRange("A1:C3").SetStyle(headerStyle);

            // Add a ListObject (table) that covers the data range A1:C6
            int tableIndex = sheet.ListObjects.Add("A1", "C6", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Convert the table to a normal range while keeping existing formatting
            table.ConvertToRange();

            // Prepare ODS save options
            OdsSaveOptions odsOptions = new OdsSaveOptions();

            // Save the workbook as ODS
            workbook.Save("TableConvertedToRange.ods", odsOptions);
        }
    }
}
