// Title: How to set built‑in Author and Title properties and a custom Application property in an Excel workbook, then verify them with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add custom document property named Application to an .xlsx file | C# read built‑in Author and Title properties from a saved workbook using Aspose.Cells | verify that a custom Application property persists after saving and reloading an Excel workbook with Aspose.Cells .NET
// Tags: add custom document property Aspose.Cells .xlsx | set built‑in workbook properties Aspose.Cells C# | read workbook metadata after save Aspose.Cells | validate Excel file properties .NET | custom Application property Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, assigns the Author and Title built‑in properties, adds a custom 'Application' property, saves the file, reloads it, reads back the Author, Title, and Application values, and confirms that all metadata fields are present.
class Program
{
    static void Main()
    {
        // Path to the workbook file
        string filePath = "output.xlsx";

        try
        {
            // -------------------------------------------------
            // Create a new workbook and set document metadata
            // -------------------------------------------------
            Workbook wb = new Workbook();

            // Set built‑in document properties (metadata)
            wb.BuiltInDocumentProperties.Author = "John Doe";
            wb.BuiltInDocumentProperties.Title = "Sample Workbook";

            // Set custom property for application name (Built‑in 'Application' not available in this version)
            wb.CustomDocumentProperties.Add("Application", "MyAsposeApp");

            // Save the workbook to disk
            wb.Save(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating or saving workbook: {ex.Message}");
            return;
        }

        // -------------------------------------------------
        // Load the saved workbook and verify metadata fields
        // -------------------------------------------------
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            Workbook loadedWb = new Workbook(filePath);

            // Retrieve built‑in properties
            string author = loadedWb.BuiltInDocumentProperties.Author;
            string title = loadedWb.BuiltInDocumentProperties.Title;

            // Retrieve custom 'Application' property
            string application = string.Empty;
            if (loadedWb.CustomDocumentProperties.Contains("Application"))
            {
                var prop = loadedWb.CustomDocumentProperties["Application"];
                application = prop?.Value?.ToString() ?? string.Empty;
            }

            // Output the results
            Console.WriteLine("Author: " + (string.IsNullOrEmpty(author) ? "Missing" : author));
            Console.WriteLine("Title: " + (string.IsNullOrEmpty(title) ? "Missing" : title));
            Console.WriteLine("Application: " + (string.IsNullOrEmpty(application) ? "Missing" : application));

            // Simple validation
            bool allPresent = !string.IsNullOrEmpty(author) &&
                              !string.IsNullOrEmpty(title) &&
                              !string.IsNullOrEmpty(application);

            Console.WriteLine(allPresent
                ? "All application metadata fields are present in the saved file."
                : "One or more metadata fields are missing.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading or processing workbook: {ex.Message}");
        }
    }
}
