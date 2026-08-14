// Title: C# – Disable “Number Stored as Text” Error Check with Aspose.Cells ErrorCheckOptions
// Description: Creates an in‑memory workbook, accesses the first worksheet’s ErrorCheckOptionCollection, adds an ErrorCheckOption, disables the NumberStoredAsText check, optionally applies it to the used range, and saves the file as NumbersAsTextErrorCheckDisabled.xlsx.
// Keywords: Aspose.Cells | ErrorCheckOptions | NumberStoredAsText | disable error check | C# | .NET | worksheet error checking | suppress green triangle | Excel warning suppression
// Common Searches: disable number stored as text warning Aspose.Cells .NET | turn off NumbersAsText error check for a worksheet | Aspose.Cells ErrorCheckOptionCollection example | suppress NumberStoredAsText error in C# | how to hide green triangle in Excel using Aspose
// Developer Intent: Turn off the NumberStoredAsText error check for a worksheet or a specific range using Aspose.Cells ErrorCheckOptions in C#.
// Use Cases: Remove the green triangle indicator after programmatically populating numeric data. | Apply the disabled check only to the used area while keeping other error checks active. | Combine disabling of NumbersStoredAsText with other error types before distributing a workbook.
// AI Prompts: Generate C# code that disables the NumberStoredAsText error check for the entire worksheet using Aspose.Cells. | Show how to disable multiple error checks, including NumbersStoredAsText, on a selected cell range with Aspose.Cells. | Explain the role of ErrorCheckOptionCollection and how to add ranges after configuring error checks in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsErrorCheckDemo
{
    // Creates an in‑memory workbook, accesses the first worksheet’s ErrorCheckOptionCollection, adds an ErrorCheckOption, disables the NumberStoredAsText check, optionally applies it to the used range, and saves the file as NumbersAsTextErrorCheckDisabled.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (in-memory)
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Access the ErrorCheckOptionCollection of the worksheet
                ErrorCheckOptionCollection errorCheckOptions = worksheet.ErrorCheckOptions;

                // Add a new ErrorCheckOption to the collection
                int optionIndex = errorCheckOptions.Add();
                ErrorCheckOption errorCheckOption = errorCheckOptions[optionIndex];

                // Disable the "Number stored as text" error check
                errorCheckOption.SetErrorCheck(ErrorCheckType.NumberStoredAsText, false);

                // Apply this setting to the used range of the worksheet (if any)
                int maxRow = worksheet.Cells.MaxRow;
                int maxCol = worksheet.Cells.MaxDataColumn;
                if (maxRow >= 0 && maxCol >= 0)
                {
                    CellArea fullRange = CellArea.CreateCellArea(0, 0, maxRow, maxCol);
                    errorCheckOption.AddRange(fullRange);
                }

                // Save the workbook to a file
                workbook.Save("NumbersAsTextErrorCheckDisabled.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
