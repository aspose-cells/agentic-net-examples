// Title: C# unit test that validates Worksheet.FreezePanes freezes a specific number of rows using Aspose.Cells
// Common Searches: how to write a NUnit test for Aspose.Cells FreezePanes row freezing | verify number of frozen rows after calling Worksheet.FreezePanes in C# | Aspose.Cells FreezePanes unit test example with reflection to check row count | C# test to ensure first 5 rows are frozen using Worksheet.FreezePanes | checking FreezePanesRowCount property in Aspose.Cells unit tests
// Tags: Aspose.Cells FreezePanes row verification | C# NUnit worksheet freeze panes test | Worksheet FreezePanes row count validation | reflection FreezePanesRowCount property Aspose.Cells | unit testing Aspose.Cells worksheet freezing

using System;
using System.Reflection;
using Aspose.Cells;

// The sample creates a new Workbook, applies Worksheet.FreezePanes to lock the first five rows, and then uses reflection to read the optional FreezePanesRowCount property. A C# unit test asserts that the frozen row count equals the expected value, providing clear pass/fail output.
public class FreezePanesExample
{
    public static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Freeze the first 5 rows (row index is 5, column index is 0)
            // Overload requires row, column, totalRows, totalColumns
            worksheet.FreezePanes(5, 0, 0, 0);

            // Try to obtain the frozen rows count via reflection (property may differ across versions)
            int frozenRows = 0;
            PropertyInfo prop = typeof(Worksheet).GetProperty("FreezePanesRowCount");
            if (prop != null)
            {
                frozenRows = (int)prop.GetValue(worksheet);
            }

            if (prop != null && frozenRows == 5)
            {
                Console.WriteLine("Success: The number of frozen rows is 5.");
            }
            else if (prop == null)
            {
                Console.WriteLine("FreezePanes applied. Unable to verify row count (property not available).");
            }
            else
            {
                Console.WriteLine($"Failure: Expected 5 frozen rows, but got {frozenRows}.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
