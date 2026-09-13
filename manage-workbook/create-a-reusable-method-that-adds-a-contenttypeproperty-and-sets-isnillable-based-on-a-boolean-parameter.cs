// Title: Create a reusable C# helper to add a custom content‑type property to an Aspose.Cells worksheet (nillable flag placeholder)
// AI Prompts: Write a static method that validates its arguments, adds a custom document property to the worksheet's workbook, and accepts a boolean parameter for a future nillable flag implementation. | Demonstrate how to instantiate a Workbook, call the helper to add a property, and save the workbook to a specified file path.
// Common Searches: asp.net add custom content type property to worksheet using Aspose.Cells | c# Aspose.Cells custom document property with optional nillable parameter | how to build a reusable Excel helper for adding custom properties in .NET | fallback to CustomDocumentProperties when CustomXmlMappings are not supported Aspose.Cells
// Tags: add custom document property Aspose.Cells | custom content type property worksheet | nillable flag placeholder Aspose.Cells | reusable Excel helper .NET | fallback CustomDocumentProperties Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Provides a static ExcelHelper.AddContentTypeProperty method that validates inputs, adds a custom document property to the workbook as a fallback for content‑type properties, accepts a boolean for a future nillable flag, and includes a sample program that creates a workbook, uses the helper, and saves the file.
public static class ExcelHelper
{
    /// <param name="worksheet">Target worksheet.</param>
    /// <param name="name">Name of the custom property.</param>
    /// <param name="value">Value of the custom property.</param>
    /// <param name="isNillable">If true, the property will be marked as nillable (not used in this fallback implementation).</param>
    public static void AddContentTypeProperty(Worksheet worksheet, string name, string value, bool isNillable)
    {
        if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
        if (string.IsNullOrEmpty(name)) throw new ArgumentException("Property name cannot be null or empty.", nameof(name));

        try
        {
            // Aspose.Cells versions prior to supporting CustomXmlMappings use CustomDocumentProperties.
            // Add a custom document property as a fallback for the content‑type property.
            var workbook = worksheet.Workbook;
            workbook.CustomDocumentProperties.Add(name, value);
        }
        catch (Exception ex)
        {
            // Wrap any exception to provide context.
            throw new InvalidOperationException("Failed to add content type property.", ex);
        }
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook (in‑memory, no template file required).
            var workbook = new Workbook();

            // Get the first worksheet.
            var worksheet = workbook.Worksheets[0];

            // Add a custom content type property.
            ExcelHelper.AddContentTypeProperty(worksheet, "MyProperty", "SomeValue", true);

            // Define output path.
            string outputPath = "output.xlsx";

            // Ensure the directory exists.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
