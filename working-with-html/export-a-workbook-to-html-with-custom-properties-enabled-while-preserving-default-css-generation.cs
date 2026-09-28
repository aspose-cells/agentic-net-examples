// Title: Export an Excel workbook to HTML with Aspose.Cells while preserving generated CSS and adding custom document properties
// AI Prompts: Create C# code that loads an existing XLSX file (or creates a new workbook), adds a custom document property, and saves the workbook as HTML using Aspose.Cells HtmlSaveOptions. | Demonstrate how to configure HtmlSaveOptions so that the default CSS stylesheet produced during HTML export is retained. | Write a try‑catch example that adds a custom property to a workbook before converting it to HTML with Aspose.Cells.
// Common Searches: how to export Excel to HTML with Aspose.Cells and keep the default CSS styles | Aspose.Cells C# add custom document property before saving as HTML | preserve generated CSS when using HtmlSaveOptions in Aspose.Cells | save workbook as HTML with custom properties using Aspose.Cells C# example
// Tags: Aspose.Cells HtmlSaveOptions CSS preservation | C# export workbook to HTML Aspose.Cells | add custom document property Aspose.Cells | HTML export default stylesheet Aspose.Cells | Excel to HTML conversion Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;

// // Loads an existing XLSX file or creates a new workbook, adds a custom document property, configures HtmlSaveOptions to keep the automatically generated CSS, and saves the workbook as an HTML file using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Load existing workbook if present; otherwise create a new empty workbook
            string inputPath = "input.xlsx";
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Add a custom document property (optional demonstration)
            workbook.CustomDocumentProperties.Add("MyCustomProperty", "CustomValue");

            // Configure HTML save options (no need to set SaveFormat; HtmlSaveOptions implies HTML)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as HTML
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
