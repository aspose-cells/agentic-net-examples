// Title: Detect circular reference errors in workbook formulas and export to ODS with Aspose.Cells for .NET
// Common Searches: how to detect circular references in formulas using Aspose.Cells C# before saving | Aspose.Cells .NET check workbook for circular formula errors and export to ODS | C# code to handle circular reference exceptions when saving a workbook as ODS with Aspose.Cells | validate formulas for circular references in Aspose.Cells and generate ODS file
// Tags: circular reference detection Aspose.Cells .NET | export workbook to ODS format Aspose.Cells | formula error handling Aspose.Cells C# | save workbook with formula validation Aspose.Cells | ODS file generation with circular reference check | Aspose.Cells workbook circular reference handling

using System;
using Aspose.Cells;

// The example creates a workbook, inserts mutually dependent formulas that form a circular reference, then attempts to save the workbook as an ODS file. The try‑catch block captures any Aspose.Cells exception raised by circular‑reference detection, demonstrating how to validate formulas and export the workbook to ODS in C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add formulas that create a circular reference (for demonstration)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Formula = "=A2";
            sheet.Cells["A2"].Formula = "=A1";

            // Save the workbook in ODS format
            string outputFile = "CircularReferenceCheck.ods";
            workbook.Save(outputFile, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
