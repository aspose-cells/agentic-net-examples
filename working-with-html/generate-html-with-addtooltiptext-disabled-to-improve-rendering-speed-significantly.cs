// Title: Aspose.Cells .NET – Export Workbook to HTML with AddTooltipText Disabled for Faster Rendering
// Description: Demonstrates how to create a workbook, insert long text, narrow a column, and save it as HTML using HtmlSaveOptions with AddTooltipText set to false. Disabling tooltips reduces file size and speeds up HTML rendering in web applications.
// Keywords: Aspose.Cells HTML export | AddTooltipText false | disable tooltip Aspose.Cells | HTML rendering performance .NET | save workbook as HTML without tooltips | Aspose.Cells HtmlSaveOptions example | C# Excel to HTML conversion
// Common Searches: Aspose.Cells disable tooltip when saving to HTML | How to improve HTML export speed with Aspose.Cells | HtmlSaveOptions AddTooltipText property usage | C# export Excel to HTML without hover tooltips | Fast HTML conversion using Aspose.Cells .NET
// Developer Intent: Turn off tooltip generation during HTML export to accelerate rendering and lower output size.
// Use Cases: Generate lightweight HTML reports from large spreadsheets. | Create web‑ready Excel views where cell overflow should not show hover text. | Optimize performance of services that repeatedly convert Excel files to HTML.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as HTML with AddTooltipText disabled and explain why this improves performance. | Show how to toggle AddTooltipText in HtmlSaveOptions based on a configuration flag in a .NET application.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, insert long text, narrow a column, and save it as HTML using HtmlSaveOptions with AddTooltipText set to false. Disabling tooltips reduces file size and speeds up HTML rendering in web applications.
    public class HtmlSaveOptionsAddTooltipTextDisabledDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data that would normally require a tooltip if it overflows
            worksheet.Cells["A1"].PutValue("This is a very long text that exceeds the column width and would normally show a tooltip.");
            worksheet.Cells.SetColumnWidth(0, 10); // Set a narrow column width to force overflow

            // Create HTML save options and disable tooltip text for better performance
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                AddTooltipText = false // Disable tooltip generation
            };

            // Save the workbook as HTML with the specified options
            string outputPath = "output_without_tooltip.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to '{outputPath}' with AddTooltipText disabled.");
        }
    }
}
