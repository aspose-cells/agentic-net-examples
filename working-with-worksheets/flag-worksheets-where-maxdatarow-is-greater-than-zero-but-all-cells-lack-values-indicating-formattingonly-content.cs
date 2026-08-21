// Title: Flag formatting‑only worksheets (MaxDataRow > 0, no values) using Aspose.Cells for .NET
// Description: Loads an Excel workbook, iterates through each worksheet, and when MaxDataRow is greater than zero while every row is blank, writes a note to cell A1 and saves the file. Shows how to detect sheets that contain only styles without data in C# with Aspose.Cells.
// Keywords: Aspose.Cells | .NET | C# | detect formatting only worksheet | MaxDataRow | empty data rows | flag sheet | Excel automation | worksheet validation | style‑only sheet
// Common Searches: Aspose.Cells detect sheet with only formatting | C# check if Excel worksheet has data | MaxDataRow no values Aspose | flag empty Excel sheets using Aspose | write note to A1 when sheet has no data
// Developer Intent: Identify worksheets that have rows but no cell values and mark them, indicating they consist solely of formatting.
// Use Cases: Automated quality‑check of incoming Excel files to flag sheets that contain only styles. | Pre‑processing step that labels formatting‑only worksheets before downstream data extraction. | Generating a report of sheets lacking data to alert users or trigger cleanup actions.
// AI Prompts: Generate C# code with Aspose.Cells that scans a workbook and adds a comment to cell A1 of any worksheet that has rows but no data values. | Suggest an alternative method using MaxDataRow or other properties to detect formatting‑only worksheets and flag them. | Explain how to modify the sample to log the names of formatting‑only sheets instead of writing to cell A1.

using System;
using Aspose.Cells;

// Loads an Excel workbook, iterates through each worksheet, and when MaxDataRow is greater than zero while every row is blank, writes a note to cell A1 and saves the file. Shows how to detect sheets that contain only styles without data in C# with Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            Cells cells = worksheet.Cells;

            // Check if there are any cells that contain data or style
            if (cells.MaxRow > 0)
            {
                bool allRowsBlank = true;

                // Examine each row up to the maximum row index
                for (int rowIndex = 0; rowIndex <= cells.MaxRow; rowIndex++)
                {
                    Row row = cells.Rows[rowIndex];

                    // Row.IsBlank is true when the row has no data (values)
                    if (!row.IsBlank)
                    {
                        allRowsBlank = false;
                        break;
                    }
                }

                // If MaxRow > 0 but every row is blank, the sheet has only formatting
                if (allRowsBlank)
                {
                    // Flag the worksheet by writing a note in cell A1
                    cells["A1"].PutValue("Formatting‑only sheet (no data)");
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
