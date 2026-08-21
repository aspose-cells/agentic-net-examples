// Title: C# – Verify Core Document Properties Remain Unchanged After Adding Custom Properties with Aspose.Cells
// Description: Demonstrates how to set built‑in properties (Author, Title) in a new Workbook, add custom document properties, save and reload the file, and confirm that the original core properties are preserved while the custom ones are persisted.
// Keywords: Aspose.Cells C# | verify built‑in properties | core document properties | custom document properties | Workbook property validation | Aspose.Cells .NET example | preserve core properties
// Common Searches: Aspose.Cells keep original author after adding custom properties | check built‑in workbook properties after save .NET | verify core properties unchanged Aspose.Cells | C# Aspose.Cells compare original and loaded document properties
// Developer Intent: Ensure that adding custom document properties does not modify the workbook's existing built‑in (core) properties.
// Use Cases: Set Author and Title, add custom fields, then programmatically assert that Author and Title are identical after reloading the workbook. | Iterate the CustomDocumentProperties collection after load to display each custom property's name and value, confirming persistence. | Integrate the boolean checks (authorUnchanged, titleUnchanged) into automated unit tests for document metadata integrity.
// AI Prompts: Generate C# code using Aspose.Cells that adds custom document properties while preserving existing Author and Title, then validates the core properties after reopening the file. | Write a method that loads a workbook from a path and returns true only if the Author and Title match expected values after custom properties have been added. | Provide a logging example that reports the verification result for core properties and lists all custom properties after a workbook is saved and reopened with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Demonstrates how to set built‑in properties (Author, Title) in a new Workbook, add custom document properties, save and reload the file, and confirm that the original core properties are preserved while the custom ones are persisted.
class VerifyCoreProperties
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set initial built‑in document properties
        workbook.BuiltInDocumentProperties["Author"].Value = "Original Author";
        workbook.BuiltInDocumentProperties["Title"].Value = "Original Title";

        // Store original values for later comparison
        object originalAuthor = workbook.BuiltInDocumentProperties["Author"].Value;
        object originalTitle = workbook.BuiltInDocumentProperties["Title"].Value;

        // Add custom document properties
        workbook.CustomDocumentProperties.Add("Project", "Alpha");
        workbook.CustomDocumentProperties.Add("Revision", 2);
        workbook.CustomDocumentProperties.Add("Approved", true);
        workbook.CustomDocumentProperties.Add("CreatedOn", DateTime.Now);

        // Save the workbook to disk
        string filePath = "VerifyCoreProperties.xlsx";
        workbook.Save(filePath);

        // Load the workbook back
        Workbook loadedWorkbook = new Workbook(filePath);

        // Verify that built‑in properties have not changed
        bool authorUnchanged = loadedWorkbook.BuiltInDocumentProperties["Author"].Value.Equals(originalAuthor);
        bool titleUnchanged = loadedWorkbook.BuiltInDocumentProperties["Title"].Value.Equals(originalTitle);

        Console.WriteLine($"Author unchanged: {authorUnchanged}");
        Console.WriteLine($"Title unchanged: {titleUnchanged}");

        // Display the custom properties to confirm they were added
        Console.WriteLine("Custom Document Properties:");
        foreach (DocumentProperty prop in loadedWorkbook.CustomDocumentProperties)
        {
            Console.WriteLine($"{prop.Name}: {prop.Value}");
        }
    }
}
