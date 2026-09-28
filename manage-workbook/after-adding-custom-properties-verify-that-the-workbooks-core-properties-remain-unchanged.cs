// Title: Add custom document properties to an Aspose.Cells workbook and ensure built‑in core properties stay unchanged (C#)
// AI Prompts: Generate C# code that creates a new Workbook, sets the Author and Title built‑in properties, adds custom properties (e.g., Project, Version, Reviewed), then checks that the Author and Title values are still the original ones. | Write a C# snippet using Aspose.Cells to compare built‑in document properties before and after inserting custom properties and output true/false results for each property.
// Common Searches: Aspose.Cells C# add custom document property without changing author property | How to verify Excel built‑in properties remain the same after adding custom properties in .NET | C# Aspose.Cells preserve workbook title when adding custom document properties | Check if built‑in document properties are unchanged after adding custom properties with Aspose.Cells
// Tags: add custom document properties Aspose.Cells C# | ensure core document properties unchanged Aspose.Cells | compare built‑in and custom properties Excel C# | verify author and title after adding custom properties | Aspose.Cells custom vs built‑in property handling

using Aspose.Cells;
using System;

// Demonstrates creating a workbook, setting built‑in Author and Title, adding custom properties (Project, Version, Reviewed), then confirming the built‑in properties remain unchanged before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set initial core (built‑in) properties
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Title = "Sample Workbook";

            // Capture original core property values for later comparison
            string originalAuthor = workbook.BuiltInDocumentProperties.Author;
            string originalTitle = workbook.BuiltInDocumentProperties.Title;

            // Add custom properties to the workbook
            workbook.CustomDocumentProperties.Add("Project", "Aspose Integration");
            workbook.CustomDocumentProperties.Add("Version", 1.0);
            workbook.CustomDocumentProperties.Add("Reviewed", true);

            // Verify that core properties have not changed after adding custom properties
            bool authorUnchanged = workbook.BuiltInDocumentProperties.Author == originalAuthor;
            bool titleUnchanged = workbook.BuiltInDocumentProperties.Title == originalTitle;

            Console.WriteLine($"Author unchanged: {authorUnchanged}");
            Console.WriteLine($"Title unchanged: {titleUnchanged}");

            // Save the workbook (optional)
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
