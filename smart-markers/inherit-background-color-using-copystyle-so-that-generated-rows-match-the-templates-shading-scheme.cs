// Title: Inherit Row Background Color with Aspose.Cells CopyStyle/CopySettings in C#
// Description: Load a template workbook, insert new rows, and copy the source row's full formatting—including background shading—using Row.CopySettings (or CopyStyle) so the generated rows match the template's shading scheme. Save the result as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | CopyStyle | CopySettings | row background color | row shading | insert rows | template formatting | smart markers
// Common Searches: Aspose.Cells copy row background color C# | CopyStyle vs CopySettings Aspose.Cells | preserve row shading after inserting rows Aspose.Cells | inherit template row style Aspose.Cells .NET | how to copy row formatting with Aspose.Cells
// Developer Intent: Copy the complete style of a template row—including its background color—to newly inserted rows using Aspose.Cells in C#.
// Use Cases: Generating reports where new data rows must follow the alternating‑color pattern defined in a template sheet. | Creating invoices that add line‑item rows while keeping the original row shading for readability. | Duplicating header or subtotal rows after insertion to maintain consistent visual formatting across a worksheet.
// AI Prompts: Provide C# code that inserts rows and copies all formatting, including background color, from a template row using Aspose.Cells Row.CopySettings with the true flag. | Explain when to use Row.CopyStyle versus Row.CopySettings in Aspose.Cells for preserving row shading. | Give a step‑by‑step tutorial for inheriting row background color from a template when generating dynamic rows with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsRowStyleCopy
{
    // Load a template workbook, insert new rows, and copy the source row's full formatting—including background shading—using Row.CopySettings (or CopyStyle) so the generated rows match the template's shading scheme. Save the result as an Excel file.
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains the desired row shading
            Workbook workbook = new Workbook("Template.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            // Source row whose style (including background color) will be copied
            Row sourceRow = worksheet.Cells.Rows[0]; // assume first row has the shading scheme

            // Insert three new rows at the desired position (e.g., after row 4)
            int insertPosition = 5; // zero‑based index where new rows start
            worksheet.Cells.InsertRows(insertPosition, 3);

            // Apply the source row's settings to each newly inserted row
            for (int i = 0; i < 3; i++)
            {
                Row targetRow = worksheet.Cells.Rows[insertPosition + i];
                // copy settings including style; true checks and gathers style when workbooks differ
                targetRow.CopySettings(sourceRow, true);
            }

            // Save the modified workbook
            workbook.Save("Result.xlsx");
        }
    }
}
