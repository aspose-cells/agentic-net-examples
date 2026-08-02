// Title: Enable Numbers Stored as Text error checking on worksheets containing numeric data – Aspose.Cells for .NET
// Description: Loads a workbook, scans each worksheet’s used range for numeric values, adds a NumbersStoredAsText error‑check option to the whole used area when numbers are present, logs the action, and saves the updated file.
// Keywords: Aspose.Cells | Numbers Stored as Text | error check | numeric data detection | C# worksheet iteration | apply error check per sheet
// Common Searches: Aspose.Cells enable Numbers Stored as Text error check per worksheet | C# detect numeric cells and set error check in Aspose.Cells | log error check changes Aspose.Cells | apply error checking only on sheets with numbers Aspose.Cells
// Developer Intent: Add the NumbersStoredAsText error check to each worksheet that contains at least one numeric cell and record which sheets were modified.
// Use Cases: Flag sheets that may have numbers stored as text before exporting to CSV. | Activate data‑validation warnings only on sheets that actually hold numeric values. | Create a console audit report of worksheets where the NumbersStoredAsText check was applied.
// AI Prompts: Generate C# code with Aspose.Cells that iterates all worksheets, detects numeric cells, enables NumbersStoredAsText error checking for those sheets, and logs the actions. | Show a LINQ‑based version that determines if a worksheet contains numeric data before applying the NumbersStoredAsText error check. | Explain how to exclude header rows from the error‑check range while still enabling NumbersStoredAsText for numeric data in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsErrorCheckDemo
{
    // Loads a workbook, scans each worksheet’s used range for numeric values, adds a NumbersStoredAsText error‑check option to the whole used area when numbers are present, logs the action, and saves the updated file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                bool hasNumeric = false;

                // Determine the used range of the sheet
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Scan cells for at least one numeric value
                for (int row = 0; row <= maxRow && !hasNumeric; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell != null && cell.Type == CellValueType.IsNumeric)
                        {
                            hasNumeric = true;
                            break;
                        }
                    }
                }

                // If numeric data is present, enable "Number Stored As Text" error check
                if (hasNumeric)
                {
                    // Access the error‑check collection for the worksheet
                    ErrorCheckOptionCollection options = sheet.ErrorCheckOptions;

                    // Add a new error‑check option
                    int optionIdx = options.Add();
                    ErrorCheckOption option = options[optionIdx];

                    // Enable checking for numbers stored as text
                    option.SetErrorCheck(ErrorCheckType.NumberStoredAsText, true);

                    // Apply the option to the whole used range of the sheet
                    CellArea usedArea = CellArea.CreateCellArea(0, 0, maxRow, maxCol);
                    option.AddRange(usedArea);

                    // Log the change
                    Console.WriteLine($"Enabled NumbersAsText error check on sheet \"{sheet.Name}\".");
                }
                else
                {
                    Console.WriteLine($"No numeric data found in sheet \"{sheet.Name}\"; error check not applied.");
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
