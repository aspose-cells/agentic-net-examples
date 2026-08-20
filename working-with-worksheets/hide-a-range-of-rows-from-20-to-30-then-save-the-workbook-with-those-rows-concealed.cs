// Title: C# Aspose.Cells Example: Hide Rows 20‑30 and Save Workbook
// Description: Shows how to hide rows 20 through 30 (zero‑based indices 19‑29) in an Aspose.Cells worksheet and save the workbook as an .xlsx file using C#.
// Keywords: Aspose.Cells hide rows C# | HideRows method Aspose.Cells | C# hide Excel rows | Aspose.Cells save hidden rows | Aspose.Cells .NET example | Excel row concealment C# | GitHub Aspose.Cells hide rows | Aspose.Cells workbook export | C# Excel automation hide rows | Aspose.Cells documentation hide rows
// Common Searches: How to hide rows 20-30 using Aspose.Cells C# | Aspose.Cells HideRows example .NET | C# code to conceal specific rows in Excel | Save Excel file with hidden rows Aspose.Cells | Aspose.Cells hide multiple rows tutorial
// Developer Intent: Hide rows 20 to 30 in a worksheet and save the workbook with those rows concealed.
// Use Cases: Create a report where intermediate calculation rows (20‑30) are hidden before distribution. | Prepare a template that hides placeholder rows so end users only see the relevant sections. | Generate a clean‑looking workbook by concealing helper rows used for internal formulas.
// AI Prompts: Write C# code using Aspose.Cells to hide rows 20‑30 and save the workbook as HiddenRows20to30.xlsx. | Explain the zero‑based indexing of the HideRows method in Aspose.Cells and how to calculate the row count. | Provide a reusable function that hides rows based on variable start index and count, then saves the file.

using System;
using Aspose.Cells;

namespace HideRowsExample
{
    // Shows how to hide rows 20 through 30 (zero‑based indices 19‑29) in an Aspose.Cells worksheet and save the workbook as an .xlsx file using C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Optionally populate some data (not required for hiding)
            for (int i = 0; i < 40; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
            }

            // Hide rows 20 to 30 (zero‑based index: 19 to 29, total 11 rows)
            worksheet.Cells.HideRows(19, 11);

            // Save the workbook with hidden rows concealed
            workbook.Save("HiddenRows20to30.xlsx");
        }
    }
}
