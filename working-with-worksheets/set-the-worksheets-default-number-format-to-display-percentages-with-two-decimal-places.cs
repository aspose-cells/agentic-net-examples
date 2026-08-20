// Title: Set Worksheet Default Percentage Format (0.00%) with Aspose.Cells for .NET
// Description: Creates a new Workbook, defines a Style with the built‑in 0.00% format, and applies it to a range that covers the entire worksheet using a StyleFlag that targets only the number format. A sample cell demonstrates the result, and the file is saved as an Excel workbook.
// Keywords: Aspose.Cells | C# | default number format | percentage style | two decimal places | apply to whole sheet | Style | StyleFlag | Excel workbook | range formatting
// Common Searches: Aspose.Cells set default percentage format for worksheet | apply 0.00% number format to all cells .NET | how to format entire sheet as percentage Aspose.Cells | C# Aspose.Cells style flag number format only | set worksheet default number format without looping
// Developer Intent: Apply a two‑decimal percentage format as the default for every cell in a worksheet.
// Use Cases: Define a Style with Number = 10 (0.00%) and use a StyleFlag to change only the number format across the sheet. | Create a large range (e.g., 0,0 to 1000 rows × 100 columns) and apply the style in a single call, avoiding per‑cell iteration. | Insert a decimal value (e.g., 0.456) into any cell to verify it displays as 45.60% after the default format is set. | Save the workbook so the percentage formatting persists in the generated Excel file.
// AI Prompts: Write C# code with Aspose.Cells that sets the worksheet's default number format to 0.00% and applies it to the entire sheet. | Explain how Style and StyleFlag can be used to modify only the number format of all cells in an Aspose.Cells worksheet. | Suggest an alternative approach to set a worksheet's default percentage format without creating a large range.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, defines a Style with the built‑in 0.00% format, and applies it to a range that covers the entire worksheet using a StyleFlag that targets only the number format. A sample cell demonstrates the result, and the file is saved as an Excel workbook.
    class SetWorksheetDefaultPercentageFormat
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a style that uses the built‑in percentage format with two decimal places (Number = 10)
                Style percentStyle = workbook.CreateStyle();
                percentStyle.Number = 10; // 0.00%

                // Prepare a StyleFlag to indicate that only the number format should be applied
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true;

                // Apply the style to a large range that effectively covers the whole worksheet.
                int totalRows = 1000;   // adjust as needed
                int totalCols = 100;    // adjust as needed

                // Use fully qualified Aspose.Cells.Range to avoid conflict with System.Range
                Aspose.Cells.Range range = worksheet.Cells.CreateRange(0, 0, totalRows, totalCols);
                range.ApplyStyle(percentStyle, flag);

                // Example: put a value in a cell to see the formatting in action
                worksheet.Cells["B2"].PutValue(0.456); // will display as 45.60%

                // Save the workbook
                workbook.Save("WorksheetDefaultPercentageFormat.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
