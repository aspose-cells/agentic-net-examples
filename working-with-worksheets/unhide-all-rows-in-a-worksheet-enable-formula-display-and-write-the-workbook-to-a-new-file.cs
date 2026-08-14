// Title: C# Example: Unhide All Rows, Show Formulas, and Save Workbook with Aspose.Cells
// Description: Demonstrates how to use Aspose.Cells for .NET to unhide every row in a worksheet, enable the ShowFormulas flag so cells display their formulas, and write the modified workbook to a new Excel file. The sample also shows how to calculate the row range using MaxDataRow and apply auto‑fit height.
// Keywords: Aspose.Cells C# unhide rows | ShowFormulas Aspose.Cells | save workbook new file Aspose.Cells | Aspose.Cells UnhideRows example | .NET Excel row visibility | Aspose.Cells GitHub sample | Excel debugging Aspose.Cells
// Common Searches: How to unhide all rows in an Aspose.Cells worksheet C# | Enable formula view (ShowFormulas) with Aspose.Cells .NET | Save modified workbook as new file using Aspose.Cells | Aspose.Cells UnhideRows MaxDataRow usage | Aspose.Cells example on GitHub for row visibility
// Developer Intent: Unhide every row, turn on formula display, and save the workbook as a new file.
// Use Cases: Restore hidden rows after temporary processing before exporting the workbook. | Display formulas for auditing while keeping original values unchanged. | Create a debug copy of an Excel file with all rows visible and formulas shown. | Prepare a clean version of a report for reviewers who need to see calculation logic.
// AI Prompts: Write C# code using Aspose.Cells to unhide all rows, enable ShowFormulas, and save the workbook to a new .xlsx file. | Explain how the UnhideRows method works, including the meaning of the height parameter and how to determine the row count with MaxDataRow. | Show how to apply the unhide and ShowFormulas settings to every worksheet in a workbook with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsUnhideRowsAndShowFormulas
{
    // Demonstrates how to use Aspose.Cells for .NET to unhide every row in a worksheet, enable the ShowFormulas flag so cells display their formulas, and write the modified workbook to a new Excel file. The sample also shows how to calculate the row range using MaxDataRow and apply auto‑fit height.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Hide rows 3,4,5 (zero‑based index) to demonstrate the unhide operation
                cells.HideRows(2, 3);

                // Unhide all rows in the worksheet.
                // Use UnhideRows starting from row 0 up to the last used row.
                // Height = -1 means auto‑fit (no explicit height change).
                int lastRow = cells.MaxDataRow;               // last row that contains data
                int totalRows = lastRow + 1;                  // total rows to process (0‑based)
                cells.UnhideRows(0, totalRows, -1);

                // Enable formula display for the worksheet.
                // When ShowFormulas is true, cells will display their formulas instead of calculated values.
                worksheet.ShowFormulas = true;

                // Save the modified workbook to a new file.
                string outputPath = "UnhiddenRows_ShowFormulas.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
