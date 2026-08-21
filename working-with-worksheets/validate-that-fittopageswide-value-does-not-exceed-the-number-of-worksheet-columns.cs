// Title: Validate and Adjust FitToPagesWide Based on Worksheet Column Count – Aspose.Cells for .NET
// Description: Creates a workbook, fills several columns, uses Worksheet.Cells.MaxColumn to get the actual column count, compares it with a desired FitToPagesWide value, caps the setting to the column count, optionally sets FitToPagesTall to 0, and saves the workbook.
// Keywords: Aspose.Cells | FitToPagesWide | MaxColumn | page setup | column count validation | .NET | C# | Excel printing layout | page scaling | worksheet columns
// Common Searches: Aspose.Cells validate FitToPagesWide against column count | how to limit FitToPagesWide to actual columns in C# | set FitToPagesWide dynamically using MaxColumn | prevent FitToPagesWide overflow in Aspose.Cells | adjust page setup width based on data columns
// Developer Intent: Ensure the FitToPagesWide property never exceeds the number of columns that contain data.
// Use Cases: Avoid truncated prints by capping FitToPagesWide to the real column count before exporting. | Automatically adapt page layout for reports with variable column numbers. | Provide a reusable helper that validates and sets FitToPagesWide for any worksheet in a multi‑sheet workbook.
// AI Prompts: Generate a C# utility method that receives a Worksheet and a desired FitToPagesWide, checks Worksheet.Cells.MaxColumn, and sets the appropriate PageSetup values. | Create code that logs a warning when the requested FitToPagesWide exceeds the column count and then applies the maximum allowed value using Aspose.Cells. | Write an example that iterates through all worksheets in a workbook, validates each worksheet's FitToPagesWide against its column count, and updates the page setup accordingly.

using System;
using Aspose.Cells;

namespace FitToPagesWideValidation
{
    // Creates a workbook, fills several columns, uses Worksheet.Cells.MaxColumn to get the actual column count, compares it with a desired FitToPagesWide value, caps the setting to the column count, optionally sets FitToPagesTall to 0, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data across several columns
            for (int col = 0; col < 20; col++)
            {
                sheet.Cells[0, col].PutValue($"Header {col + 1}");
                sheet.Cells[1, col].PutValue($"Data {col + 1}");
            }

            // Desired FitToPagesWide value (example)
            int desiredFitToPagesWide = 25;

            // Determine the total number of columns that contain data
            // MaxColumn is zero‑based, so add 1 to get the count
            int totalColumns = sheet.Cells.MaxColumn + 1;

            // Validate the FitToPagesWide setting
            if (desiredFitToPagesWide > totalColumns)
            {
                Console.WriteLine($"FitToPagesWide ({desiredFitToPagesWide}) exceeds the number of columns ({totalColumns}). Adjusting to {totalColumns}.");
                sheet.PageSetup.FitToPagesWide = totalColumns;
            }
            else
            {
                sheet.PageSetup.FitToPagesWide = desiredFitToPagesWide;
            }

            // Optionally set FitToPagesTall to 0 to let height adjust automatically
            sheet.PageSetup.FitToPagesTall = 0;

            // Save the workbook
            workbook.Save("FitToPagesWideValidated.xlsx");
        }
    }
}
