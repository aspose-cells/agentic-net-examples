// Title: Export an Aspose.Cells workbook to HTML with default CSS and custom document properties in C#
// AI Prompts: Write C# code that adds a custom document property to a Workbook and saves it as HTML using Aspose.Cells with the built‑in HtmlSaveOptions. | Show how to use HtmlSaveOptions to export an Excel workbook to HTML while keeping the automatically generated CSS and retaining any custom properties.
// Common Searches: C# Aspose.Cells export workbook to HTML preserving generated CSS | How to include custom document properties when saving Excel as HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions default styling example in C# | Save Excel file to HTML with custom properties using Aspose.Cells library | Export workbook to HTML with default CSS and custom properties Aspose.Cells C#
// Tags: Aspose.Cells HtmlSaveOptions default CSS | export workbook to HTML Aspose.Cells | custom property handling Aspose.Cells | C# Excel to HTML conversion Aspose.Cells | preserve generated CSS Aspose.Cells HTML export

using System;
using Aspose.Cells;

// The sample creates a workbook, adds a custom document property, and saves it as an HTML file using Aspose.Cells with default HtmlSaveOptions, which retains the automatically generated CSS.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate some data (optional, just for demonstration)
            workbook.Worksheets[0].Cells["A1"].PutValue("Sample Text");

            // Add a custom document property
            workbook.CustomDocumentProperties.Add("MyCustomProperty", "CustomValue");

            // Configure HTML save options (default settings are sufficient)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Export the workbook to HTML with the specified options
            string outputPath = "ExportedWorkbook.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
