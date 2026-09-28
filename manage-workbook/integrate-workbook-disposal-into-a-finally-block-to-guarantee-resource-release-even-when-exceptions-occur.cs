// Title: Guarantee Aspose.Cells Workbook disposal with a finally block in C#
// AI Prompts: Wrap the Workbook creation, operations, and Save call in a try block and add a finally block that checks for null before calling Dispose(). | Rewrite the example using a C# using statement so the Aspose.Cells Workbook is automatically disposed. | Add error logging in the catch block while still ensuring the Workbook is released in every execution path.
// Common Searches: c# ensure aspose.cells workbook is disposed in finally even on error | how to release aspose.cells resources safely after saving workbook | using statement vs try/finally for disposing Aspose.Cells Workbook in .NET | example of exception‑safe workbook cleanup with Aspose.Cells | best practice for disposing Aspose.Cells Workbook in console application
// Tags: Aspose.Cells workbook disposal | finally block resource cleanup C# | using statement Aspose.Cells | exception‑safe workbook release | C# Aspose.Cells resource management | dispose workbook after Save

using System;
using Aspose.Cells;

// // Creates an Aspose.Cells Workbook, writes a value to cell A1, saves it as Result.xlsx, and guarantees the workbook is disposed in a finally block so resources are released even if an exception occurs.
class Program
{
    static void Main()
    {
        Workbook workbook = null;

        try
        {
            // {CreateWorkbook}
            workbook = new Workbook();

            // Example operation: write a value to the first worksheet
            workbook.Worksheets[0].Cells["A1"].PutValue("Hello Aspose!");

            // {SaveWorkbook}
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during processing
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Guarantee that the workbook resources are released
            if (workbook != null)
                workbook.Dispose();
        }
    }
}
