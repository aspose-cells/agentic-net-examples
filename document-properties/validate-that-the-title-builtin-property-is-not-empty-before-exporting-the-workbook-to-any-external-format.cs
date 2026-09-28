// Title: Ensure the Title built-in document property is set before exporting an Aspose.Cells workbook to PDF in C#
// AI Prompts: Verify that workbook.BuiltInDocumentProperties.Title contains a non-empty string and raise an InvalidOperationException if it does not, then save the workbook as PDF. | Add a pre-export check for the Title built-in property in an Aspose.Cells workbook and handle missing titles gracefully in C#. | Create a helper method that validates the Title document property and calls Workbook.Save with SaveFormat.Pdf only when the validation passes.
// Common Searches: c# Aspose.Cells how to check Title built-in property before saving to PDF | throw error when Excel workbook title property is empty using Aspose.Cells | validate document properties in Aspose.Cells before export | ensure non-empty Title property in Aspose.Cells workbook C# example
// Tags: title property check Aspose.Cells | pre-export document property validation C# | InvalidOperationException for missing title Aspose.Cells | save workbook as PDF after property verification | Aspose.Cells built-in document property handling

using System;
using Aspose.Cells;

// Creates a workbook, sets the Title built-in document property, validates that the Title is not empty, and saves the workbook as a PDF, throwing an InvalidOperationException if the Title is missing.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Set the Title built‑in property to satisfy validation
            workbook.BuiltInDocumentProperties.Title = "Sample Report";

            // Validate that the Title built‑in property is not empty
            if (string.IsNullOrWhiteSpace(workbook.BuiltInDocumentProperties.Title))
            {
                throw new InvalidOperationException("The Title built‑in property must not be empty before exporting.");
            }

            // Export the workbook to PDF
            workbook.Save("ExportedReport.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
