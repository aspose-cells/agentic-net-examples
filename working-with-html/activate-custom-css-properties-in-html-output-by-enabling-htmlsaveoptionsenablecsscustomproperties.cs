// Title: C# – Export Aspose.Cells Workbook to HTML Using CSS Custom Properties
// Description: Demonstrates how to create a workbook, apply formatting, and save it as HTML with HtmlSaveOptions.EnableCssCustomProperties set to true, producing lean CSS‑based styling.
// Keywords: Aspose.Cells HTML export | EnableCssCustomProperties | C# HtmlSaveOptions | CSS custom properties Aspose | Excel to HTML with variables | Aspose.Cells .NET example
// Common Searches: Aspose.Cells enable CSS variables in HTML output | HtmlSaveOptions.EnableCssCustomProperties C# example | How to reduce duplicate styles when exporting Excel to HTML | Export workbook to HTML with CSS custom properties Aspose
// Developer Intent: Activate CSS custom properties for the HTML generated from an Aspose.Cells workbook.
// Use Cases: Generate lightweight HTML reports from Excel data with reusable style definitions. | Apply cell formatting once and reuse it across multiple elements via CSS variables. | Include images once and reference them through custom properties to minimize HTML size.
// AI Prompts: Show how to turn off CSS custom properties in Aspose.Cells HTML export. | Provide a C# snippet that assigns a custom CSS class to a range and saves with EnableCssCustomProperties enabled. | Explain the performance impact of using HtmlSaveOptions.EnableCssCustomProperties for large workbooks.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomCssDemo
{
    // Demonstrates how to create a workbook, apply formatting, and save it as HTML with HtmlSaveOptions.EnableCssCustomProperties set to true, producing lean CSS‑based styling.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data with formatting
            sheet.Cells["A1"].PutValue("Hello World");
            var style = sheet.Cells["A1"].GetStyle();
            style.Font.IsBold = true;
            style.Font.Color = System.Drawing.Color.Blue;
            sheet.Cells["A1"].SetStyle(style);

            // Insert an image that will be reused (optional, demonstrates benefit of CSS custom properties)
            // Ensure you have a valid image file path or use a placeholder image.
            // int imgIdx = sheet.Pictures.Add(2, 2, "logo.png");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable CSS custom properties to optimize the HTML output
            htmlOptions.EnableCssCustomProperties = true;

            // Save the workbook as HTML with the custom CSS property enabled (lifecycle: save)
            workbook.Save("OutputWithCssCustomProperties.html", htmlOptions);

            Console.WriteLine("HTML file saved with EnableCssCustomProperties = true.");
        }
    }
}
