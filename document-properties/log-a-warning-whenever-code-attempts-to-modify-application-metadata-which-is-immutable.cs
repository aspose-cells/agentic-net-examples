// Title: Log a warning when trying to change the immutable NameOfApplication property in Aspose.Cells for .NET
// Description: This C# example creates a Workbook, registers a custom IWarningCallback to capture Aspose.Cells warnings, attempts to set BuiltInDocumentProperties.NameOfApplication (which is read‑only), logs a pre‑emptive warning, catches any exception caused by the immutability, and finally saves the file.
// Keywords: Aspose.Cells .NET | NameOfApplication immutable | log warning Aspose.Cells | IWarningCallback example | modify built‑in document properties | C# workbook metadata | prevent application metadata change
// Common Searches: Aspose.Cells log warning when changing NameOfApplication | Can I modify BuiltInDocumentProperties.NameOfApplication in .NET | IWarningCallback usage for immutable properties Aspose.Cells | How to capture Aspose.Cells warnings for document metadata | Immutable application metadata Aspose.Cells C#
// Developer Intent: Show how to detect and warn developers when code attempts to modify the read‑only NameOfApplication property of an Aspose.Cells workbook.
// Use Cases: Implement a custom IWarningCallback that prints warning type and related object for any Aspose.Cells warning. | Attempt to assign workbook.BuiltInDocumentProperties.NameOfApplication, catch the resulting exception, and log an informative error message. | Save the workbook after the modification attempt while ensuring all warnings are recorded.
// AI Prompts: Generate a C# IWarningCallback that logs Aspose.Cells warnings when a developer tries to change an immutable built‑in document property. | Explain how to programmatically check if NameOfApplication is read‑only and emit a warning before assignment. | Write unit tests that verify a warning is raised and an exception is caught when setting workbook.BuiltInDocumentProperties.NameOfApplication.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// This C# example creates a Workbook, registers a custom IWarningCallback to capture Aspose.Cells warnings, attempts to set BuiltInDocumentProperties.NameOfApplication (which is read‑only), logs a pre‑emptive warning, catches any exception caused by the immutability, and finally saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Register a warning callback to capture any Aspose.Cells warnings
        workbook.Settings.WarningCallback = new MyWarningHandler();

        // Attempt to modify the application metadata (NameOfApplication)
        // This metadata is considered immutable; we log a warning before the attempt
        try
        {
            Console.WriteLine("Attempting to modify application metadata (NameOfApplication)...");
            Console.WriteLine("Warning: Application metadata is immutable. Modification may be ignored or cause an error.");

            // Perform the modification (if allowed)
            workbook.BuiltInDocumentProperties.NameOfApplication = "MyCustomApp";
        }
        catch (Exception ex)
        {
            // Log any exception that occurs due to immutability
            Console.WriteLine($"Error while modifying application metadata: {ex.Message}");
        }

        // Save the workbook (optional)
        workbook.Save("output.xlsx");
    }

    // Implementation of the warning callback interface
    class MyWarningHandler : IWarningCallback
    {
        public void Warning(WarningInfo warningInfo)
        {
            // Log the warning type and related object
            Console.WriteLine($"Aspose.Cells Warning: {warningInfo.WarningType}");
            if (warningInfo.ErrorObject != null)
            {
                Console.WriteLine($"Related Object: {warningInfo.ErrorObject}");
            }
        }
    }
}
