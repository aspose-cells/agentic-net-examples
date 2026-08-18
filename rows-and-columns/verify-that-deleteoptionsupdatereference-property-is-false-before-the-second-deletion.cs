// Title: Aspose.Cells .NET – Verify DeleteOptions.UpdateReference Is False Before Second Row Deletion
// Description: C# example that creates a workbook, fills column A, deletes the second row with DeleteOptions.UpdateReference = true, then confirms DeleteOptions.UpdateReference defaults to false before deleting another row without updating references, and saves the file.
// Keywords: Aspose.Cells DeleteOptions | UpdateReference false | C# delete rows Aspose.Cells | verify DeleteOptions property | row deletion without formula update | Aspose.Cells .NET example
// Common Searches: Aspose.Cells DeleteOptions.UpdateReference default value | how to delete rows without updating formulas in Aspose.Cells | check DeleteOptions.UpdateReference before DeleteRows | multiple row deletions with different UpdateReference settings | C# Aspose.Cells delete rows example
// Developer Intent: Confirm that DeleteOptions.UpdateReference is false before executing the second DeleteRows call to avoid unintended reference updates.
// Use Cases: Delete a row while shifting formula references, then delete another row while preserving existing references. | Implement a data‑cleanup routine that selectively updates references on the first deletion only. | Programmatically validate DeleteOptions settings at runtime to prevent accidental formula changes.
// AI Prompts: Write C# code using Aspose.Cells that deletes one row with UpdateReference = true and a second row with UpdateReference = false, including runtime checks of the property. | Explain how DeleteOptions.UpdateReference influences formula references across worksheets in Aspose.Cells and show how to verify its value before calling DeleteRows. | Provide a step‑by‑step guide to test the default value of DeleteOptions.UpdateReference and ensure it remains false for subsequent DeleteRows operations.

using System;
using Aspose.Cells;

// C# example that creates a workbook, fills column A, deletes the second row with DeleteOptions.UpdateReference = true, then confirms DeleteOptions.UpdateReference defaults to false before deleting another row without updating references, and saves the file.
public class VerifyDeleteOptions
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data in column A (rows 0-4)
            for (int i = 0; i < 5; i++)
            {
                cells[i, 0].PutValue(i + 1); // A1..A5 = 1..5
            }

            // ---------- First deletion ----------
            // Create DeleteOptions with UpdateReference set to true
            DeleteOptions firstOptions = new DeleteOptions
            {
                UpdateReference = true
            };

            // Delete the second row (index 1) using the options
            // This will update references in other worksheets (if any)
            cells.DeleteRows(1, 1, firstOptions);

            // Verify that the property is indeed true (optional check)
            if (firstOptions.UpdateReference)
            {
                Console.WriteLine("First deletion: UpdateReference is true as expected.");
            }

            // ---------- Second deletion ----------
            // Create DeleteOptions without setting UpdateReference (defaults to false)
            DeleteOptions secondOptions = new DeleteOptions();

            // Verify that UpdateReference is false before performing the deletion
            if (!secondOptions.UpdateReference)
            {
                Console.WriteLine("Second deletion: UpdateReference is false as expected.");

                // Delete the (original) fourth row (now at index 2 after previous deletion)
                cells.DeleteRows(2, 1, secondOptions);
            }
            else
            {
                // This block should not be reached; included for completeness
                Console.WriteLine("Unexpected: UpdateReference is true before second deletion.");
            }

            // Save the workbook to verify the result
            workbook.Save("VerifyDeleteOptions.xlsx");
            Console.WriteLine("Workbook saved as VerifyDeleteOptions.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            VerifyDeleteOptions.Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
