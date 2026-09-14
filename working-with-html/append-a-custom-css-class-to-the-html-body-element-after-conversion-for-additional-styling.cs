// Title: Append a custom CSS class to the <body> tag of HTML generated from an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that converts an Aspose.Cells Workbook to HTML and injects a specified CSS class into the <body> element of the resulting markup. | Show how to read the HTML output from Aspose.Cells, apply a regular expression to add or merge a class attribute on the <body> tag, and write the modified file back to disk.
// Common Searches: how to add a CSS class to the body tag after saving a workbook as HTML with Aspose.Cells | Aspose.Cells .NET insert custom class into generated HTML body element | C# regex replace body tag in HTML produced by Aspose.Cells | modify Aspose.Cells HTML output to include custom stylesheet class
// Tags: Aspose.Cells HTML body class injection | C# regex modify Aspose.Cells generated HTML | add custom CSS class during Excel to HTML conversion | Aspose.Cells HTML save options custom styling | custom-body class Aspose.Cells output

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example creates a workbook, adds data, saves it as HTML with Aspose.Cells using specific save options, reads the HTML into a string, uses a regular expression to append a "custom-body" CSS class to the <body> tag, writes the modified HTML to a file, and logs a completion message.
class HtmlConversionWithCustomCss
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // empty workbook
        // Example: add some data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Set HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportImagesAsBase64 = true, // embed images
            ExportActiveWorksheetOnly = true // export only the first sheet
        };

        // Save the workbook to a memory stream as HTML
        using (MemoryStream htmlStream = new MemoryStream())
        {
            workbook.Save(htmlStream, saveOptions);
            htmlStream.Position = 0;

            // Read the generated HTML into a string
            string htmlContent;
            using (StreamReader reader = new StreamReader(htmlStream, Encoding.UTF8))
            {
                htmlContent = reader.ReadToEnd();
            }

            // Append a custom CSS class to the <body> tag
            // This simple replace works when <body> has no existing attributes
            // For a more robust solution, a regex can handle existing attributes
            string pattern = @"<body([^>]*)>";
            string replacement = "<body$1 class=\"custom-body\">";
            string modifiedHtml = Regex.Replace(htmlContent, pattern, replacement, RegexOptions.IgnoreCase);

            // Optionally, write the modified HTML to a file
            File.WriteAllText("output.html", modifiedHtml, Encoding.UTF8);
        }

        Console.WriteLine("HTML conversion completed with custom CSS class added to the body element.");
    }
}
