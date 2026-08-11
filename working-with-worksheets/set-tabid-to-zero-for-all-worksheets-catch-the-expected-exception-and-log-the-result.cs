// Title: C# – Set Worksheet TabId to Zero with Aspose.Cells, Catch CellsException and Log Results
// Description: Demonstrates how to create a workbook, add worksheets, attempt to assign a TabId of 0 to each sheet, handle the expected CellsException, log success or error details, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | TabId | CellsException | worksheet exception handling | log TabId assignment | save workbook | invalid TabId | Excel API | Aspose.Cells .NET example
// Common Searches: Aspose.Cells set TabId to zero | C# catch CellsException when setting TabId | how to handle invalid TabId in Aspose.Cells | log worksheet TabId errors Aspose | save workbook after TabId change Aspose.Cells
// Developer Intent: Assign a TabId of 0 to every worksheet, capture the expected CellsException, and record the outcome.
// Use Cases: Verify that setting an illegal TabId triggers a CellsException. | Provide detailed logging for both successful and failed TabId assignments. | Ensure the workbook is saved regardless of exception occurrences.
// AI Prompts: Generate C# code using Aspose.Cells to set TabId = 0 for each worksheet, catch CellsException, and log results. | Explain why a TabId of 0 causes a CellsException in Aspose.Cells and outline best practices for handling it.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, add worksheets, attempt to assign a TabId of 0 to each sheet, handle the expected CellsException, log success or error details, and save the file using Aspose.Cells for .NET.
class SetTabIdDemo
{
    static void Main()
    {
        // Create a new workbook (contains one default worksheet)
        Workbook workbook = new Workbook();

        // Add additional worksheets for demonstration
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Attempt to set TabId to zero for each worksheet and log the outcome
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            try
            {
                sheet.TabId = 0; // This may throw an exception
                Console.WriteLine($"TabId set to 0 for worksheet '{sheet.Name}'.");
            }
            catch (CellsException ex)
            {
                // Expected exception handling
                Console.WriteLine($"Error setting TabId for worksheet '{sheet.Name}': {ex.Message}");
                Console.WriteLine($"Exception Type: {ex.Code}");
            }
            catch (Exception ex)
            {
                // Catch any other unexpected exceptions
                Console.WriteLine($"Unexpected error on worksheet '{sheet.Name}': {ex.Message}");
            }
        }

        // Save the workbook to persist any changes
        string outputPath = "TabIdDemo.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
