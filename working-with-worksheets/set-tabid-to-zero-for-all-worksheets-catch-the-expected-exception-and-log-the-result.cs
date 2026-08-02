// Title: Aspose.Cells C# Example: Set Worksheet TabId to Zero, Catch CellsException, and Log Results
// Description: Demonstrates how to create a Workbook, add worksheets, attempt to assign TabId = 0 to each sheet, capture the expected CellsException, output success or error messages to the console, and save the file.
// Keywords: Aspose.Cells | C# | Worksheet TabId | set TabId zero | CellsException | exception handling | invalid TabId | console logging | save workbook | GitHub example
// Common Searches: Aspose.Cells set TabId to zero | What exception is thrown when TabId is zero in Aspose.Cells | C# catch CellsException for worksheet TabId | Iterate worksheets and modify TabId Aspose.Cells | Log TabId assignment failures Aspose.Cells
// Developer Intent: Set every worksheet's TabId to zero, handle the expected CellsException, and record the outcome.
// Use Cases: Validate that TabId cannot be zero and identify affected worksheets. | Show proper error handling for worksheet property changes in a new workbook. | Generate a console report of success or failure before saving the workbook.
// AI Prompts: Create a reusable method that sets TabId to zero for all worksheets, catches CellsException, and returns a list of results with worksheet name, success flag, and error details. | Rewrite the example to log TabId assignment errors to a text file instead of the console. | Explain why Aspose.Cells throws a CellsException when TabId is set to zero and recommend best practices for handling such validation errors.

using System;
using Aspose.Cells;

// Demonstrates how to create a Workbook, add worksheets, attempt to assign TabId = 0 to each sheet, capture the expected CellsException, output success or error messages to the console, and save the file.
class SetTabIdDemo
{
    static void Main()
    {
        // Create a new workbook and add a couple of worksheets for demonstration
        Workbook workbook = new Workbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Iterate through all worksheets and attempt to set TabId to zero
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            try
            {
                sheet.TabId = 0;
                Console.WriteLine($"Worksheet '{sheet.Name}': TabId set to {sheet.TabId} successfully.");
            }
            catch (CellsException ex) // Expected Aspose.Cells specific exception
            {
                Console.WriteLine($"Worksheet '{sheet.Name}': Failed to set TabId. Exception: {ex.Message}, Code: {ex.Code}");
            }
            catch (Exception ex) // Any other unexpected exception
            {
                Console.WriteLine($"Worksheet '{sheet.Name}': Unexpected error: {ex.Message}");
            }
        }

        // Save the workbook (optional, just to complete the lifecycle)
        workbook.Save("TabIdDemo.xlsx");
    }
}
