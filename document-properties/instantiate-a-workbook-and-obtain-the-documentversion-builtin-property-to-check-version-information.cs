// Title: Read the DocumentVersion built‑in property from an Aspose.Cells workbook in C#
// Common Searches: C# Aspose.Cells get DocumentVersion from BuiltInDocumentProperties | how to read Excel file version using Aspose.Cells API in .NET | access built‑in document properties of a workbook with Aspose.Cells | retrieve workbook version information programmatically with Aspose.Cells
// Tags: Aspose.Cells BuiltInDocumentProperties API | read DocumentVersion property C# | retrieve Excel workbook version Aspose.Cells | access built‑in document properties .NET | get workbook version information Aspose.Cells

using Aspose.Cells;
using System;

// The example creates an empty Workbook, accesses its BuiltInDocumentProperties collection, safely reads the "DocumentVersion" property (defaulting to "Unknown" if missing), and prints the version to the console, with exception handling for unexpected errors.
class Program
{
    static void Main()
    {
        try
        {
            // Instantiate a new (empty) workbook
            Workbook workbook = new Workbook();

            // Access the built‑in document properties collection
            var builtInProps = workbook.BuiltInDocumentProperties;

            // Retrieve the "DocumentVersion" property value safely
            string documentVersion = builtInProps["DocumentVersion"]?.Value?.ToString() ?? "Unknown";

            // Display the version information
            Console.WriteLine("Document Version: " + documentVersion);
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
