// Title: Load an Excel workbook and add a custom document property "ClientName" only if it does not already exist using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that opens an existing .xlsx file, checks the CustomDocumentProperties collection for a property named 'ClientName', and adds it when missing. | Show how to iterate over workbook.CustomDocumentProperties in Aspose.Cells to avoid creating duplicate custom properties before saving the file. | Provide a C# example that includes comprehensive error handling for loading, modifying, and saving an Excel workbook while safely managing custom document properties.
// Common Searches: aspnet check if custom document property exists before adding with Aspose.Cells | c# Aspose.Cells add custom property only if not present | how to avoid duplicate custom properties in an Excel workbook using Aspose.Cells | load workbook and manage custom document properties in .NET
// Tags: conditional addition of custom document property Aspose.Cells | verify custom property existence in Excel C# | prevent duplicate custom property Aspose.Cells | load workbook and manage custom properties .NET | exception handling for Aspose.Cells workbook operations

using Aspose.Cells;
using System;
using System.IO;

// The program loads an existing Excel workbook, inspects its CustomDocumentProperties for a property named "ClientName", adds the property with a sample value only if it is absent, and then saves the workbook to a new file, handling missing files and other exceptions gracefully.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the existing file
            workbook = new Workbook(inputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Access the collection of custom document properties
        var customProps = workbook.CustomDocumentProperties;

        // Check whether a property named "ClientName" already exists
        bool clientNameExists = false;
        foreach (var prop in customProps)
        {
            // The collection items expose Name property
            if (prop.Name == "ClientName")
            {
                clientNameExists = true;
                break;
            }
        }

        // Add the "ClientName" property only if it does not exist
        if (!clientNameExists)
        {
            try
            {
                // Example value; replace with actual client name as needed
                customProps.Add("ClientName", "Acme Corp");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to add custom property: {ex.Message}");
                // Continue; saving the workbook without the new property is still possible
            }
        }

        // Save the workbook to a new file (or overwrite the original)
        try
        {
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
