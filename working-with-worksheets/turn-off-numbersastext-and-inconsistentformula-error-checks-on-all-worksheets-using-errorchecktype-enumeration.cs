// Title: Aspose.Cells .NET – Turn Off NumberStoredAsText & InconsistentFormula Checks for Every Worksheet
// Description: Learn how to disable the NumberStoredAsText and InconsistentFormula error checks on all worksheets in a workbook using Aspose.Cells for .NET. The example creates (or loads) a Workbook, iterates through each Worksheet, adds an ErrorCheckOption, disables the two checks, applies the option to the used cell range only when data exists, and saves the file.
// Keywords: Aspose.Cells disable NumberStoredAsText | Aspose.Cells turn off InconsistentFormula | ErrorCheckOptionCollection .NET | disable worksheet error checks Aspose | set error check options all sheets | C# Aspose.Cells error checking | Excel error checks suppression
// Common Searches: how to disable NumberStoredAsText error in Aspose.Cells | turn off InconsistentFormula check for all worksheets .NET | Aspose.Cells disable specific error checks workbook | programmatically suppress Excel error warnings with Aspose | apply error‑check settings to used range Aspose.Cells
// Developer Intent: Disable the NumberStoredAsText and InconsistentFormula error checks on every worksheet in an Aspose.Cells workbook.
// Use Cases: Generate reports that contain numeric strings without triggering NumberStoredAsText warnings. | Copy formulas that intentionally differ across rows while keeping the error list clean. | Apply the settings only to worksheets that contain data, leaving empty sheets untouched.
// AI Prompts: Write C# code using Aspose.Cells to turn off NumberStoredAsText and InconsistentFormula checks for all worksheets and save the workbook. | Show how to modify the example to apply the error‑check options to a custom cell range instead of the entire used area. | Demonstrate how to re‑enable the NumberStoredAsText and InconsistentFormula checks later in the same workbook.

using Aspose.Cells;
using System;

// Learn how to disable the NumberStoredAsText and InconsistentFormula error checks on all worksheets in a workbook using Aspose.Cells for .NET. The example creates (or loads) a Workbook, iterates through each Worksheet, adds an ErrorCheckOption, disables the two checks, applies the option to the used cell range only when data exists, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (you can also load an existing file here)
            Workbook workbook = new Workbook();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Get the collection that holds error‑check options for the worksheet
                ErrorCheckOptionCollection errorCheckOptions = worksheet.ErrorCheckOptions;

                // Add a new ErrorCheckOption to the collection
                int optionIndex = errorCheckOptions.Add();
                ErrorCheckOption errorCheckOption = errorCheckOptions[optionIndex];

                // Turn off the "Number stored as text" check
                errorCheckOption.SetErrorCheck(ErrorCheckType.NumberStoredAsText, false);
                // Turn off the "Inconsistent formula" check
                errorCheckOption.SetErrorCheck(ErrorCheckType.InconsistFormula, false);

                // Determine the used range; MaxDataRow/Column return -1 if the sheet is empty
                int maxRow = worksheet.Cells.MaxDataRow;
                int maxColumn = worksheet.Cells.MaxDataColumn;

                // Apply the settings only when there is at least one used cell
                if (maxRow >= 0 && maxColumn >= 0)
                {
                    CellArea fullRange = CellArea.CreateCellArea(0, 0, maxRow, maxColumn);
                    errorCheckOption.AddRange(fullRange);
                }
            }

            // Save the workbook with the updated error‑check settings
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
