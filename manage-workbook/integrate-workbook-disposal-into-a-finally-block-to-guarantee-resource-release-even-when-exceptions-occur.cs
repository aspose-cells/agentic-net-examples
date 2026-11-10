// Title: Dispose an Aspose.Cells Workbook in a C# finally block to guarantee resource release after saving
// Common Searches: C# how to ensure Aspose.Cells workbook is disposed even when an exception occurs | using finally block to release Aspose.Cells resources after saving Excel file | best practice for disposing Aspose.Cells Workbook in .NET applications | exception-safe cleanup of Aspose.Cells workbook in C# code | guarantee resource release for Aspose.Cells workbook with try‑catch‑finally
// Tags: Aspose.Cells workbook disposal pattern | C# finally block resource cleanup | exception-safe Aspose.Cells handling | dispose workbook after saving Excel | resource release for Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a Workbook, adds sample data, saves it as Result.xlsx, and uses a finally block to call Dispose on the Workbook, ensuring that all Excel-related resources are released even if an exception is thrown.
class Program
{
    static void Main()
    {
        Workbook workbook = null;
        try
        {
            // Create a new workbook
            workbook = new Workbook();

            // Populate the workbook with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["B1"].PutValue(12345);

            // Save the workbook to a file
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during processing
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Guarantee that the workbook is disposed and resources are released
            if (workbook != null)
            {
                workbook.Dispose();
            }
        }
    }
}
