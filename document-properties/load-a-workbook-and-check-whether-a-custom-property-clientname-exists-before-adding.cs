// Title: Load an Excel workbook and add a custom 'ClientName' property only if it does not already exist using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check if a custom document property exists before adding | how to add a custom Excel property conditionally with Aspose.Cells | prevent duplicate custom properties when updating an .xlsx file in C# | verify presence of 'ClientName' custom property in a workbook using Aspose.Cells
// Tags: conditional addition of custom document property Aspose.Cells | custom property existence verification Excel .NET | add client name custom property workbook C# | load and save workbook with modified properties Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, checks the CustomDocumentProperties collection for a "ClientName" entry, adds the property with a value only when it is missing, and then saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Access the custom document properties collection
            var customProps = workbook.CustomDocumentProperties;

            // Add the "ClientName" property only if it does not already exist
            if (!customProps.Contains("ClientName"))
            {
                // Replace with the actual client name as needed
                customProps.Add("ClientName", "Acme Corp");
            }

            // Save the workbook (overwrites or creates a new file)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
