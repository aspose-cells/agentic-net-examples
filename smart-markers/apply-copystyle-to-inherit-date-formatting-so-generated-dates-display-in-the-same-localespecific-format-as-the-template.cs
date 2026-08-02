// Title: Copy date formatting from a template workbook using Aspose.Cells for .NET
// Description: Demonstrates how to load a template Excel file, extract the date style from a cell, copy that style with Style.Copy, and apply it to newly generated DateTime values so the output respects the template's locale‑specific date format.
// Keywords: Aspose.Cells | .NET | C# | Style.Copy | date format inheritance | Excel template style | locale specific dates | copy cell style | Excel automation
// Common Searches: Aspose.Cells copy date style from template | inherit Excel date format in C# | apply template cell style to new dates Aspose | preserve regional date format when generating Excel | Style.Copy example for date cells
// Developer Intent: Copy a date‑format style defined in a template workbook and reuse it for dates generated in a new workbook.
// Use Cases: Populate a financial report while keeping the date format defined by the corporate template. | Generate invoices that automatically match the regional date format set in a master workbook. | Create a scheduling spreadsheet that inherits locale‑specific date formatting from an existing Excel file.
// AI Prompts: Show me C# code that copies a date style from a template workbook and applies it to multiple date cells using Aspose.Cells. | Explain how to use Style.Copy to preserve locale‑specific date formatting when generating Excel files with Aspose.Cells for .NET. | Provide a step‑by‑step example of inheriting a date format from a template and applying it to new dates across workbooks.

using System;
using Aspose.Cells;

namespace AsposeCellsCopyDateStyleDemo
{
    // Demonstrates how to load a template Excel file, extract the date style from a cell, copy that style with Style.Copy, and apply it to newly generated DateTime values so the output respects the template's locale‑specific date format.
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains the desired date format style
            Workbook templateWorkbook = new Workbook("Template.xlsx");
            Worksheet templateSheet = templateWorkbook.Worksheets[0];
            // Assume the template date style is applied to cell A1
            Cell templateDateCell = templateSheet.Cells["A1"];
            Style templateDateStyle = templateDateCell.GetStyle();

            // Create a new workbook where dates will be generated
            Workbook resultWorkbook = new Workbook();
            Worksheet resultSheet = resultWorkbook.Worksheets[0];

            // Generate some dates in the result workbook
            resultSheet.Cells["A1"].PutValue(DateTime.Now);
            resultSheet.Cells["A2"].PutValue(new DateTime(2023, 12, 25));
            resultSheet.Cells["A3"].PutValue(new DateTime(2024, 1, 1));

            // Create a new style in the result workbook and copy the template style into it
            Style copiedDateStyle = resultWorkbook.CreateStyle();
            copiedDateStyle.Copy(templateDateStyle);

            // Apply the copied style to the generated date cells
            resultSheet.Cells["A1"].SetStyle(copiedDateStyle);
            resultSheet.Cells["A2"].SetStyle(copiedDateStyle);
            resultSheet.Cells["A3"].SetStyle(copiedDateStyle);

            // Save the result workbook
            resultWorkbook.Save("Result.xlsx");
        }
    }
}
