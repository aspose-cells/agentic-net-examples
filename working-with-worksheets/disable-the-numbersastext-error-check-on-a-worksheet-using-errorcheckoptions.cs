// Title: Disable "Number Stored as Text" error check in Aspose.Cells for .NET
// Description: Demonstrates how to turn off the NumberStoredAsText warning on a worksheet using Aspose.Cells. The example creates a workbook, accesses the first sheet, retrieves the ErrorCheckOptionCollection, adds an ErrorCheckOption, disables the specific error type, defines the used cell area, applies the setting, and saves the file.
// Keywords: Aspose.Cells disable NumberStoredAsText | ErrorCheckOptionCollection SetErrorCheck false | turn off number stored as text warning | .NET Excel error checking | worksheet error check configuration
// Common Searches: how to disable number stored as text error in Aspose.Cells | Aspose.Cells C# turn off NumberStoredAsText for whole sheet | set error check options programmatically Aspose.Cells | disable Excel warning for numbers stored as text using .NET
// Developer Intent: Programmatically suppress the "Number stored as text" error indicator for a worksheet (or a specific range) with Aspose.Cells.
// Use Cases: Generate Excel reports where numeric values are intentionally stored as text without visual warnings. | Prepare data‑import templates that contain leading apostrophes and need to hide the error flag. | Apply the disabled check only to populated cells while leaving empty areas untouched.
// AI Prompts: Show C# code to disable the NumberStoredAsText error check for a selected range with Aspose.Cells. | Give an example of enabling and disabling multiple error checks (e.g., NumberStoredAsText, InconsistentFormula) on a worksheet. | Explain how to read existing error‑check settings from a workbook and modify them using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to turn off the NumberStoredAsText warning on a worksheet using Aspose.Cells. The example creates a workbook, accesses the first sheet, retrieves the ErrorCheckOptionCollection, adds an ErrorCheckOption, disables the specific error type, defines the used cell area, applies the setting, and saves the file.
    class DisableNumbersAsTextErrorCheck
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Access the collection of error‑check options for the worksheet
                ErrorCheckOptionCollection errorCheckOptions = worksheet.ErrorCheckOptions;

                // Add a new error‑check option to the collection
                int optionIndex = errorCheckOptions.Add();
                ErrorCheckOption errorCheckOption = errorCheckOptions[optionIndex];

                // Disable the "Number stored as text" error check
                errorCheckOption.SetErrorCheck(ErrorCheckType.NumberStoredAsText, false);

                // Determine the used range; if the sheet is empty, use a single cell (A1)
                int endRow = Math.Max(0, worksheet.Cells.MaxDataRow);
                int endColumn = Math.Max(0, worksheet.Cells.MaxDataColumn);
                CellArea usedArea = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = endRow,
                    EndColumn = endColumn
                };

                // Apply the setting to the determined range
                errorCheckOption.AddRange(usedArea);

                // Save the workbook
                workbook.Save("DisableNumbersAsTextErrorCheck.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
