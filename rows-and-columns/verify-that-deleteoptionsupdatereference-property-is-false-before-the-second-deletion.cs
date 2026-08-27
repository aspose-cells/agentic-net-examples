// Title: How to delete a row with UpdateReference=true and then conditionally delete another row with UpdateReference=false using Aspose.Cells for .NET
// AI Prompts: Write C# code that removes a specific row from an Aspose.Cells worksheet with DeleteOptions.UpdateReference set to true, then removes a different row only after confirming DeleteOptions.UpdateReference is false. | Demonstrate how to check the UpdateReference flag of a DeleteOptions instance before performing a second row deletion in a .NET workbook.
// Common Searches: Aspose.Cells C# delete row with UpdateReference true then false | verify DeleteOptions.UpdateReference before second row deletion .NET | conditional row removal using DeleteOptions in Aspose.Cells workbook | prevent reference updates when deleting rows with Aspose.Cells | C# example of DeleteRows with different UpdateReference settings
// Tags: DeleteRows UpdateReference true Aspose.Cells | conditional row deletion DeleteOptions .NET | verify DeleteOptions.UpdateReference false | Aspose.Cells row removal without reference update | C# workbook row deletion based on DeleteOptions flag

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, fills column A with values 1‑5, deletes row 2 using DeleteOptions with UpdateReference = true, checks that a second DeleteOptions instance has UpdateReference = false, then deletes row 3, and finally saves the workbook.
    public class VerifyDeleteOptionsUpdateReference
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data in column A (A1:A5 = 1,2,3,4,5)
                for (int i = 0; i < 5; i++)
                {
                    cells[i, 0].PutValue(i + 1);
                }

                // First deletion: delete row 2 (index 1) with UpdateReference = true
                DeleteOptions firstOptions = new DeleteOptions
                {
                    UpdateReference = true
                };
                cells.DeleteRows(1, 1, firstOptions);

                // Second deletion: delete row 3 (original index 3, now index 2) with UpdateReference = false
                DeleteOptions secondOptions = new DeleteOptions
                {
                    UpdateReference = false
                };

                if (!secondOptions.UpdateReference)
                {
                    cells.DeleteRows(2, 1, secondOptions);
                    Console.WriteLine("Second deletion performed with UpdateReference = false.");
                }
                else
                {
                    Console.WriteLine("UpdateReference is not false; aborting second deletion.");
                }

                // Save the workbook to verify the result
                string outputPath = "VerifyDeleteOptionsUpdateReference.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
