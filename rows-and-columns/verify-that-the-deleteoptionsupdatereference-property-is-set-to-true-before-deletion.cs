// Title: C# – Verify DeleteOptions.UpdateReference = true before deleting columns with Aspose.Cells
// Description: Demonstrates how to configure DeleteOptions.UpdateReference, validate the flag, delete a column, and automatically adjust formulas in an Aspose.Cells workbook using C#.
// Keywords: Aspose.Cells DeleteOptions | UpdateReference true | C# DeleteColumns formula update | validate DeleteOptions before deletion | Aspose.Cells column removal
// Common Searches: Aspose.Cells verify DeleteOptions.UpdateReference | C# delete column and keep formulas correct | DeleteOptions.UpdateReference usage example | how to update formula references after column deletion Aspose.Cells | check DeleteOptions flag before DeleteColumns
// Developer Intent: Ensure DeleteOptions.UpdateReference is enabled so that column deletion updates any dependent formulas.
// Use Cases: Prevent runtime errors by confirming the UpdateReference flag before removing columns that affect formulas. | Automatically shift formula references when a column containing source data is deleted. | Generate a workbook that reflects updated calculations after structural changes.
// AI Prompts: Write C# code that deletes a row only when DeleteOptions.UpdateReference is true using Aspose.Cells. | Show how to catch InvalidOperationException if DeleteOptions.UpdateReference is false before a delete operation. | Explain the effect of DeleteOptions.UpdateReference on formula references when columns are removed in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to configure DeleteOptions.UpdateReference, validate the flag, delete a column, and automatically adjust formulas in an Aspose.Cells workbook using C#.
    public class VerifyDeleteOptionsUpdateReference
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate some sample data
                cells["A1"].PutValue(10);
                cells["B1"].PutValue(20);
                cells["C1"].Formula = "=A1+B1"; // Formula referencing column A and B

                // Create DeleteOptions and set UpdateReference to true
                DeleteOptions deleteOptions = new DeleteOptions
                {
                    UpdateReference = true
                };

                // Verify that UpdateReference is true before performing deletion
                if (!deleteOptions.UpdateReference)
                {
                    throw new InvalidOperationException("DeleteOptions.UpdateReference must be set to true before deletion.");
                }

                // Delete column A (index 0) using the DeleteOptions
                // This will also update the formula in C1 to reference the new column positions
                worksheet.Cells.DeleteColumns(0, 1, deleteOptions);

                // Output the updated formula to confirm reference adjustment
                Console.WriteLine("Updated formula in C1 after column deletion: " + cells["C1"].Formula);

                // Save the workbook
                string outputPath = "VerifyDeleteOptionsUpdateReference_Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
