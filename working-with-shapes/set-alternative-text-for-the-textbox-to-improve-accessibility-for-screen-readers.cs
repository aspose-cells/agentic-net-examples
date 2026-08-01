// Title: Add Alt Text to a TextBox Shape in Aspose.Cells for .NET (C#) – Improve Screen‑Reader Accessibility
// Description: C# example that creates a workbook, inserts a TextBox shape, sets the AlternativeText property for screen‑reader support, optionally adds visible text, and saves the file as an .xlsx document. Demonstrates how to make Excel reports accessible with Aspose.Cells.
// Keywords: Aspose.Cells C# | Add Alt Text to TextBox | AlternativeText property | Excel shape accessibility | screen reader support | Aspose.Cells .NET example | accessibility metadata | US developers | UK developers | India developers | GitHub Aspose.Cells sample
// Common Searches: Aspose.Cells set textbox alt text C# | How to add AlternativeText to a shape in Aspose.Cells | Make Excel TextBox accessible with Aspose.Cells | C# code for textbox accessibility in Aspose.Cells | Aspose.Cells alternative text for screen readers
// Developer Intent: Assign a descriptive string to the AlternativeText property of a TextBox shape so that screen readers can convey its purpose.
// Use Cases: Provide alt text for report textboxes to meet WCAG accessibility standards. | Automatically populate AlternativeText for dynamically generated textboxes in financial dashboards. | Ensure exported PDFs retain textbox descriptions for compliant accessibility tags.
// AI Prompts: Generate C# Aspose.Cells code that creates a TextBox, sets AlternativeText and custom font, then saves the workbook. | Show how to iterate over all TextBox shapes in a worksheet and assign AlternativeText based on adjacent cell values. | Provide an example that adds AlternativeText to a TextBox and exports the workbook to PDF while preserving accessibility metadata.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, inserts a TextBox shape, sets the AlternativeText property for screen‑reader support, optionally adds visible text, and saves the file as an .xlsx document. Demonstrates how to make Excel reports accessible with Aspose.Cells.
class SetTextboxAlternativeText
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a TextBox shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.TextBox, 2, 2, 0, 0, 100, 200);

            // Cast the generic Shape to a TextBox to access TextBox‑specific members
            TextBox textBox = (TextBox)shape;

            // Set the alternative (alt) text for accessibility (screen readers)
            textBox.AlternativeText = "Summary of sales data for Q1 2024";

            // Optionally set visible text inside the textbox
            textBox.Text = "Q1 Sales Summary";

            // Save the workbook
            workbook.Save("TextboxWithAltText.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
