// Title: Turn off NumberStoredAsText and InconsistentFormula error checks for all worksheets (Aspose.Cells C#)
// Description: Loads a workbook, iterates through each worksheet, adds an ErrorCheckOption, disables the NumberStoredAsText and InconsistentFormula checks for the full used range, and saves the updated file.
// Keywords: Aspose.Cells | C# | ErrorCheckOption | ErrorCheckType | NumberStoredAsText | InconsistentFormula | disable error checks | worksheet error indicators | Excel automation | bulk workbook settings
// Common Searches: Aspose.Cells turn off NumberStoredAsText for all sheets | disable InconsistentFormula error check C# Aspose.Cells | apply error‑check options to entire workbook Aspose | remove Excel error flags before saving with Aspose.Cells | bulk disable error checks in Aspose.Cells workbook
// Developer Intent: Disable the NumberStoredAsText and InconsistentFormula error checks on every worksheet in an Aspose.Cells workbook.
// Use Cases: Prepare reports for distribution without visible Excel error icons. | Clean up programmatically generated workbooks before PDF conversion. | Enforce consistent error‑checking settings across all sheets in an automated pipeline.
// AI Prompts: Write C# code that disables NumberStoredAsText and InconsistentFormula checks for all worksheets using Aspose.Cells. | Show how to add EmptyCellReference to the list of disabled error checks in the same workbook. | Explain how to target a custom cell range instead of the entire used area when setting error‑check options.

using System;
using Aspose.Cells;

// Loads a workbook, iterates through each worksheet, adds an ErrorCheckOption, disables the NumberStoredAsText and InconsistentFormula checks for the full used range, and saves the updated file.
class DisableErrorChecks
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the collection of error‑check options for the current worksheet
            ErrorCheckOptionCollection options = sheet.ErrorCheckOptions;

            // Add a new error‑check option entry
            int optionIndex = options.Add();
            ErrorCheckOption option = options[optionIndex];

            // Disable the "Number stored as text" check
            option.SetErrorCheck(ErrorCheckType.NumberStoredAsText, false);
            // Disable the "Inconsistent formula" check
            option.SetErrorCheck(ErrorCheckType.InconsistFormula, false);

            // Apply the settings to the entire used range of the worksheet
            int maxRow = sheet.Cells.MaxRow;
            int maxCol = sheet.Cells.MaxDataColumn;
            CellArea fullArea = CellArea.CreateCellArea(0, 0, maxRow, maxCol);
            option.AddRange(fullArea);
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
