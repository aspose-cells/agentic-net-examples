// Title: Verify that Aspose.Cells throws a CellsException when accessing a cell with an illegal display name in C#
// Common Searches: Aspose.Cells C# catch CellsException for invalid cell name | how to test invalid cell display name exception in Aspose.Cells .NET | retrieving a cell by illegal name throws error Aspose.Cells | unit test for cell name validation using Aspose.Cells API | exception thrown when using sheet.Cells["Invalid!Name"] in C#
// Tags: CellsException for illegal cell names | Aspose.Cells cell reference validation | C# workbook invalid cell name test | sheet.Cells indexer exception handling | Aspose.Cells display name error handling

using System;
using Aspose.Cells;

// The example creates a workbook, accesses the first worksheet, and attempts to retrieve a cell using the illegal display name "Invalid!Name". It expects a CellsException, demonstrating proper exception handling for invalid cell display names in Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        try
        {
            // Attempt to retrieve a cell using an invalid display name.
            // The name "Invalid!Name" contains an illegal character and should trigger an exception.
            Cell invalidCell = sheet.Cells["Invalid!Name"];

            // If no exception is thrown, the test has failed.
            Console.WriteLine("Test failed: No exception was thrown for an invalid cell name.");
        }
        catch (CellsException ex)
        {
            // Expected exception type for an invalid cell name.
            Console.WriteLine("Test passed: Caught expected CellsException.");
            Console.WriteLine("Exception message: " + ex.Message);
        }
        catch (Exception ex)
        {
            // Any other exception type indicates an unexpected failure.
            Console.WriteLine("Test failed: Caught unexpected exception type: " + ex.GetType().FullName);
            Console.WriteLine("Exception message: " + ex.Message);
        }
    }
}
