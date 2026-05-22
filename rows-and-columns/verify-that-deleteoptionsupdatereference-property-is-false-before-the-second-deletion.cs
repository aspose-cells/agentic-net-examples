using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
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

                // Populate some sample data and a formula that references the data
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);
                cells["A3"].PutValue(30);
                cells["B1"].Formula = "=SUM(A1:A3)";

                // First deletion: delete the first row with UpdateReference set to true
                DeleteOptions firstOptions = new DeleteOptions
                {
                    UpdateReference = true // references will be updated after this deletion
                };
                // Delete row 0 (first row) using the overload that accepts DeleteOptions
                cells.DeleteRows(0, 1, firstOptions);

                // Prepare DeleteOptions for the second deletion with UpdateReference set to false
                DeleteOptions secondOptions = new DeleteOptions
                {
                    UpdateReference = false // we want to verify that this is false before deletion
                };

                // Verify that UpdateReference is false before performing the second deletion
                if (!secondOptions.UpdateReference)
                {
                    Console.WriteLine("DeleteOptions.UpdateReference is false before the second deletion as expected.");
                }
                else
                {
                    Console.WriteLine("Unexpected: DeleteOptions.UpdateReference is true before the second deletion.");
                }

                // Second deletion: delete the next row (now at index 0 after the first deletion)
                cells.DeleteRows(0, 1, secondOptions);

                // Save the workbook to verify the final state
                workbook.Save("VerifyDeleteOptionsUpdateReference_Output.xlsx");
                Console.WriteLine("Workbook saved successfully.");
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
            VerifyDeleteOptionsUpdateReference.Run();
        }
    }
}