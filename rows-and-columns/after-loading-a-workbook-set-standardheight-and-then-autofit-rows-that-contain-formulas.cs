// Title: C# – Set Standard Row Height and Auto‑Fit Only Formula Rows with Aspose.Cells
// Description: Load a workbook, assign a default row height, detect rows that contain at least one formula, and auto‑fit those rows using Aspose.Cells for .NET before saving the file.
// Keywords: Aspose.Cells C# set row height | StandardHeight Aspose.Cells | AutoFitRow formula rows | auto fit rows with formulas .NET | iterate cells Aspose.Cells | C# spreadsheet row height example | Aspose.Cells selective auto‑fit
// Common Searches: Aspose.Cells set default row height C# | auto fit rows containing formulas Aspose.Cells | how to use AutoFitRow for formula rows in .NET | C# Aspose.Cells iterate rows to find formulas | set StandardHeight then auto‑fit specific rows Aspose.Cells
// Developer Intent: Define a uniform row height and auto‑fit only rows that have formulas.
// Use Cases: Create reports where static rows keep a fixed height while calculated rows expand to show full results. | Prepare workbooks for printing, ensuring formula‑driven rows automatically adjust without altering other rows. | Generate financial statements where only rows with formulas need dynamic height for readability.
// AI Prompts: Generate C# code with Aspose.Cells that sets StandardHeight and auto‑fits only rows containing formulas, handling merged cells appropriately. | Explain how to modify the loop to skip hidden rows while still auto‑fitting rows that contain formulas. | Provide a commented Aspose.Cells example that sets a default row height and selectively auto‑fits rows based on formula presence.

using System;
using Aspose.Cells;

// Load a workbook, assign a default row height, detect rows that contain at least one formula, and auto‑fit those rows using Aspose.Cells for .NET before saving the file.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the default row height (in points)
        worksheet.Cells.StandardHeight = 18; // adjust as needed

        // Determine the range of used rows and columns
        int maxRow = worksheet.Cells.MaxDataRow;
        int maxCol = worksheet.Cells.MaxDataColumn;

        // Auto‑fit only rows that contain at least one formula
        for (int row = 0; row <= maxRow; row++)
        {
            bool containsFormula = false;

            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = worksheet.Cells[row, col];
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    containsFormula = true;
                    break;
                }
            }

            if (containsFormula)
            {
                // Auto‑fit the specific row
                worksheet.AutoFitRow(row);
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
