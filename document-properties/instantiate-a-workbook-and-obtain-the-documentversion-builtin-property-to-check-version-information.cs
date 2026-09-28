// Title: How to read the DocumentVersion built‑in property from a new Aspose.Cells workbook in C#
// AI Prompts: Write C# code that creates an empty Aspose.Cells Workbook and prints the value of its DocumentVersion built‑in property. | Show a C# example that verifies the existence of the DocumentVersion property in BuiltInDocumentProperties before reading it, and returns a default string when it is absent. | Demonstrate proper exception handling while retrieving the DocumentVersion property from a workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# retrieve DocumentVersion from workbook built‑in properties | how to check if DocumentVersion exists in Excel file using Aspose.Cells .NET | default value when DocumentVersion property is not set in Aspose.Cells workbook
// Tags: read DocumentVersion built‑in property Aspose.Cells | C# check built‑in document properties Aspose.Cells | fallback value for missing DocumentVersion Aspose.Cells | exception handling Aspose.Cells workbook property access | access built‑in properties of new Excel workbook C#

using Aspose.Cells;
using System;

// Creates an empty Workbook, accesses its BuiltInDocumentProperties collection, safely reads the "DocumentVersion" property if present, supplies a default string when missing, and writes the result to the console with error handling.
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

            // Retrieve the value of the "DocumentVersion" property safely
            string documentVersion = "N/A";
            if (builtInProps.Contains("DocumentVersion") && builtInProps["DocumentVersion"]?.Value != null)
            {
                documentVersion = builtInProps["DocumentVersion"].Value.ToString();
            }

            // Display the version information
            Console.WriteLine("Document Version: " + documentVersion);
        }
        catch (Exception ex)
        {
            // Runtime safety: report any unexpected errors
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
