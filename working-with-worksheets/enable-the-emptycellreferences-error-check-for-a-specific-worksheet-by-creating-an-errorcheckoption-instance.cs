// Title: Enable EmptyCellReference Error Check for a Worksheet with Aspose.Cells in C# (.NET)
// Description: Demonstrates how to create a Workbook, retrieve its ErrorCheckOptionCollection, add an ErrorCheckOption, turn on the EmptyCellRef check (green‑triangle warning), apply it to the worksheet's used range via CellArea, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | .NET error checking | EmptyCellRef | ErrorCheckOption | worksheet error check | green triangle warning | CellArea range | programmatic error check | save workbook with error check
// Common Searches: Aspose.Cells enable EmptyCellRef error check C# | how to add ErrorCheckOption to worksheet | set EmptyCellReference warning in .NET | apply error check to used range Aspose.Cells | green triangle for empty cell references
// Developer Intent: Programmatically turn on the EmptyCellReference error check for a specific worksheet and apply it to the sheet’s used range.
// Use Cases: Highlight formulas that reference empty cells in automatically generated financial reports. | Validate data integrity before sharing a workbook with end users. | Add the EmptyCellRef check to multiple worksheets during batch processing.
// AI Prompts: Generate C# code that enables EmptyCellRef error checking for all worksheets in an Aspose.Cells workbook. | Show how to disable the EmptyCellReference error check for a selected cell range using ErrorCheckOption. | Explain how to list all active error‑check types for a given worksheet with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a Workbook, retrieve its ErrorCheckOptionCollection, add an ErrorCheckOption, turn on the EmptyCellRef check (green‑triangle warning), apply it to the worksheet's used range via CellArea, and save the file using Aspose.Cells for .NET.
    public class EnableEmptyCellReferenceErrorCheck
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the collection of error‑check options for this worksheet
                ErrorCheckOptionCollection errorCheckOptions = worksheet.ErrorCheckOptions;

                // Add a new ErrorCheckOption to the collection
                int optionIndex = errorCheckOptions.Add();
                ErrorCheckOption errorCheckOption = errorCheckOptions[optionIndex];

                // Enable the EmptyCellRef error check (shows green triangle for formulas that refer to empty cells)
                errorCheckOption.SetErrorCheck(ErrorCheckType.EmptyCellRef, true);

                // Apply the option to the whole used range of the worksheet
                int maxRow = worksheet.Cells.MaxRow;
                int maxCol = worksheet.Cells.MaxDataColumn; // limit to actual data
                CellArea fullRange = CellArea.CreateCellArea(0, 0, maxRow, maxCol);
                errorCheckOption.AddRange(fullRange);

                // Save the workbook
                string outputPath = "EnableEmptyCellReferenceErrorCheck.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            EnableEmptyCellReferenceErrorCheck.Run();
        }
    }
}
