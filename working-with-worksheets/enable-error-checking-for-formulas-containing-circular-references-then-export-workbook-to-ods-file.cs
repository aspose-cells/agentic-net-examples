// Title: Detect circular reference errors with Aspose.Cells, recalculate formulas, and save the workbook as ODS in C#
// AI Prompts: Write C# code using Aspose.Cells that creates a workbook, inserts a circular formula, runs formula recalculation to capture circular reference errors, and saves the file as ODS. | Show how to catch and log exceptions thrown by Aspose.Cells when a circular reference is found during formula calculation, then export the workbook to ODS. | Provide a step‑by‑step example of enabling circular reference detection, performing formula recalculation, and using SaveFormat.Ods to write the workbook in C#.
// Common Searches: aspnet circular reference detection Aspose.Cells calculate formula ODS export | c# Aspose.Cells how to handle circular formula errors before saving as ODS | save workbook to OpenDocument Spreadsheet after formula calculation with Aspose.Cells | detect and log circular references in Excel files using Aspose.Cells .NET | example of using Workbook.CalculateFormula and SaveFormat.Ods in C#
// Tags: circular reference handling Aspose.Cells | formula recalculation error handling | save workbook to ODS Aspose.Cells | C# Aspose.Cells formula calculation | exception handling for circular formulas .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, adds a circular formula (A1 = A2, A2 = A1), triggers formula recalculation to detect the circular reference, catches any resulting exceptions, and saves the workbook as an ODS file using Aspose.Cells' SaveFormat.Ods.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Note: Circular reference detection is enabled by default during formula calculation.
            // Add a worksheet with a circular formula for demonstration
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Formula = "=A2";
            sheet.Cells["A2"].Formula = "=A1";

            // Recalculate formulas to trigger circular reference detection
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "CircularReferenceCheckOutput.ods";

            // Save the workbook to an ODS file (use the non‑obsolete enum value)
            workbook.Save(outputPath, SaveFormat.Ods);

            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
