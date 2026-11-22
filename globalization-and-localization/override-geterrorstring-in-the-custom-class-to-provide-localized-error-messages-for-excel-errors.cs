// Title: How to override GetErrorString in a custom class to localize Excel error messages using Aspose.Cells for .NET
// Common Searches: C# override GetErrorString Aspose.Cells to display error messages in French | custom GetErrorString implementation for localized #DIV/0! text in Aspose.Cells | Aspose.Cells error localization example with GetErrorString override | provide language‑specific Excel error strings in .NET using Aspose.Cells | how to customize Excel error messages per culture with Aspose.Cells API | override GetErrorString to return translated error messages for workbook calculations
// Tags: override GetErrorString Aspose.Cells | localized Excel error messages .NET | custom error string provider for workbook calculations | Aspose.Cells error localization C# | Excel formula error translation Aspose.Cells | culture‑aware error string implementation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsErrorLocalization
{
    // The example demonstrates creating a Workbook, inserting a formula that generates an error, and overriding the GetErrorString method in a custom class to return culture‑specific error text. After calculating formulas, the workbook is saved with localized error messages, illustrating how to provide localized Excel error strings in Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Example: set a formula that will generate a #DIV/0! error
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue(0);
                sheet.Cells["A2"].PutValue(0);
                sheet.Cells["B1"].Formula = "=A1/A2";

                // Calculate formulas to populate error values
                workbook.CalculateFormula();

                // Define output file path
                string outputPath = "LocalizedErrors.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
