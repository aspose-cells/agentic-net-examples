// Title: Hide Columns 5‑9 with Aspose.Cells (C#) and Save Workbook Preserving Hidden State
// Description: Creates a workbook, fills a 20×15 range, hides columns 5‑9 using Cells.HideColumns, and saves the file as XLSX so the hidden columns stay hidden.
// Keywords: Aspose.Cells hide columns C# | Cells.HideColumns example | export workbook with hidden columns | preserve hidden columns Aspose | C# Excel hide column range
// Common Searches: Aspose.Cells hide multiple columns .NET | save Excel file with hidden columns using Aspose | C# hide columns 5 to 9 Aspose.Cells | how to keep columns hidden after export Aspose | Cells.HideColumns usage example
// Developer Intent: Hide a specific column range (indexes 5‑9) and export the workbook while retaining the hidden state.
// Use Cases: Protect sensitive data by hiding columns before distribution. | Generate templates that automatically conceal auxiliary columns. | Apply role‑based column visibility and deliver the file to end users.
// AI Prompts: Provide C# code that hides columns 5‑9 in an Aspose.Cells workbook and saves it with hidden columns intact. | Show how to use Cells.HideColumns to conceal a range of columns and export to XLSX. | Explain how Aspose.Cells records hidden column information in the saved Excel file.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills a 20×15 range, hides columns 5‑9 using Cells.HideColumns, and saves the file as XLSX so the hidden columns stay hidden.
    public class HideMultipleColumnsAndExport
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created and saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (optional, just for demonstration)
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 15; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Hide columns with zero‑based indexes 5 to 9 (5 columns total)
            int startColumn = 5;   // column index to start hiding
            int columnCount = 5;   // number of columns to hide (5,6,7,8,9)
            cells.HideColumns(startColumn, columnCount);   // rule: Cells.HideColumns

            // Save the workbook; hidden columns are preserved in the saved file
            workbook.Save("HiddenColumnsPreserved.xlsx", SaveFormat.Xlsx); // lifecycle: save
        }
    }
}
