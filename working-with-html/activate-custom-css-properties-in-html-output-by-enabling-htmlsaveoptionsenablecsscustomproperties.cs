// Title: Enable CSS custom properties when exporting an Excel workbook to HTML using Aspose.Cells for .NET
// AI Prompts: Generate C# code that saves an Aspose.Cells workbook as HTML with the EnableCssCustomProperties flag turned on. | Show how to configure HtmlSaveOptions to produce HTML that uses CSS variables for styling in Aspose.Cells. | Convert a populated Excel worksheet to an HTML file while activating custom CSS properties via Aspose.Cells API.
// Common Searches: Aspose.Cells HtmlSaveOptions EnableCssCustomProperties example C# | How to export Excel to HTML with CSS variables using Aspose.Cells .NET | Enable custom CSS properties in HTML output from Aspose.Cells workbook | Saving workbook as HTML with CSS custom properties in C# Aspose.Cells | Aspose.Cells HTML export with CSS custom properties turned on
// Tags: Aspose.Cells HtmlSaveOptions CSS custom properties | export workbook to HTML with CSS variables | C# enable CSS custom properties Aspose.Cells | HTML output styling using CSS variables Aspose.Cells | save Excel as HTML using custom CSS properties

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The program creates a workbook, adds sample data, configures HtmlSaveOptions to enable CSS custom properties, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // creates an empty workbook

        // Add some data to demonstrate HTML output
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Score");
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(85);
        sheet.Cells["A3"].PutValue("Bob");
        sheet.Cells["B3"].PutValue(92);

        // Configure HTML save options to enable CSS custom properties
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.EnableCssCustomProperties = true; // activates custom CSS variables in the output

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", saveOptions);
    }
}
