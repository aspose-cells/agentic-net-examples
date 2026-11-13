// Title: Verify that Aspose.Cells throws a CellsException when accessing a cell with an invalid display name in C#
// Common Searches: Aspose.Cells C# catch CellsException for non-A1 cell reference | how to test invalid cell display name throws exception in Aspose.Cells | unit test Aspose.Cells invalid A1 style reference handling | C# code to validate error when using wrong cell name with Aspose.Cells
// Tags: Aspose.Cells invalid display name exception handling | CellsException for non-A1 reference in C# | C# workbook cell retrieval error validation | Aspose.Cells A1 style reference validation

using System;
using Aspose.Cells;

// Creates a workbook, accesses the first worksheet, then attempts to retrieve a cell using the invalid display name "InvalidName"; the code expects a CellsException and reports whether the expected exception was caught.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Attempt to retrieve a cell using an invalid display name.
            // "InvalidName" is not a valid A1-style reference, so Aspose.Cells
            // should throw a CellsException.
            Cell cell = sheet.Cells["InvalidName"];
        }
        catch (CellsException ex)
        {
            // Expected exception for an invalid cell reference.
            Console.WriteLine("Expected exception caught: " + ex.Message);
        }
        catch (Exception ex)
        {
            // Any other exception means the test failed.
            Console.WriteLine("Unexpected exception: " + ex.GetType().Name + " - " + ex.Message);
        }
    }
}
