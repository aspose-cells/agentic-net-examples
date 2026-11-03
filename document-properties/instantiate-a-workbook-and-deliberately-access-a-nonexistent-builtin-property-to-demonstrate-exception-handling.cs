// Title: C# example: handling exception when accessing a missing built‑in document property with Aspose.Cells Workbook
// Common Searches: Aspose.Cells C# catch exception for invalid BuiltInDocumentProperties key | how to safely read a non‑existent built‑in property from an Excel workbook using Aspose.Cells | example of exception handling when accessing unknown built‑in document property in .NET | retrieve built‑in document property that does not exist with Aspose.Cells and handle error | Aspose.Cells workbook built‑in property access throws KeyNotFoundException handling
// Tags: Aspose.Cells BuiltInDocumentProperties key error handling | C# missing Excel built‑in property exception | Workbook property access error Aspose.Cells | .NET Aspose.Cells document property not found handling | invalid built‑in property handling in Aspose.Cells

using System;
using Aspose.Cells;

// The code creates a new Workbook, attempts to read a built‑in document property named "NonExistentProperty", catches the resulting exception, and prints the exception message to the console, demonstrating proper error handling for missing properties.
class Program
{
    static void Main()
    {
        // Instantiate a new workbook
        Workbook workbook = new Workbook();

        try
        {
            // Deliberately access a non‑existent built‑in property
            var value = workbook.BuiltInDocumentProperties["NonExistentProperty"];
            Console.WriteLine("Property value: " + value);
        }
        catch (Exception ex)
        {
            // Handle the expected exception
            Console.WriteLine("Exception caught: " + ex.Message);
        }
    }
}
