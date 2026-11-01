// Title: Create a new Excel workbook and add a link‑to‑content custom document property "SourceUrl" using Aspose.Cells for .NET (C#)
// Common Searches: how to add a URL as a custom document property in an Excel file using Aspose.Cells C# | Aspose.Cells example for setting a link‑to‑content property named SourceUrl | store external website link in workbook custom properties with Aspose.Cells .NET | C# code to create a workbook and add a custom property containing a web address using Aspose.Cells | add link‑to‑content custom property to an Excel workbook programmatically in .NET
// Tags: custom document property with URL Aspose.Cells | link-to-content property Excel C# | add SourceUrl custom property Aspose.Cells | initialize workbook and set custom property .NET | Aspose.Cells workbook custom properties example

using System;
using Aspose.Cells;

namespace Example
{
    // The example creates a new Workbook, adds a custom document property called "SourceUrl" that holds an external URL string, saves the workbook as Output.xlsx, and prints a confirmation message.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a custom document property (stored as a string)
                workbook.CustomDocumentProperties.Add(
                    "SourceUrl",
                    "https://www.example.com");

                // Save the workbook to verify the property is stored
                string outputPath = "Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
