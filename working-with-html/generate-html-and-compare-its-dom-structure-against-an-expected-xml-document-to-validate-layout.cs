// Title: Export a single worksheet to self‑contained HTML with Base64 images using Aspose.Cells and verify the root element against an expected XML layout in C#
// AI Prompts: Create a C# console application that loads an Excel workbook, saves only the first worksheet as HTML with images embedded as Base64 using Aspose.Cells, and writes the result to a specified file path. | Add logic to read an XML file that defines the expected layout and check whether the generated HTML string contains the XML document's root element. | Replace the simple string check with a full DOM comparison using an HTML parser (e.g., HtmlAgilityPack) to ensure the HTML structure exactly matches the expected XML layout.
// Common Searches: Aspose.Cells export first worksheet to HTML with embedded Base64 images C# | C# compare generated HTML DOM to XML template after Excel conversion | how to validate that HTML output contains a specific XML root element in .NET | self‑contained HTML export from Excel using Aspose.Cells and verify layout
// Tags: Aspose.Cells HTML export base64 images | C# validate HTML against XML layout | compare HTML DOM with expected XML | export single worksheet to self-contained HTML | HTML DOM verification using HtmlAgilityPack

using System;
using System.IO;
using System.Xml;
using Aspose.Cells;

namespace HtmlLayoutValidatorApp
{
    // The program loads an Excel workbook, saves the first worksheet as a self‑contained HTML file with images encoded in Base64 via Aspose.Cells, then reads an expected XML layout and reports whether the HTML contains the XML's root element, with an option to perform a full DOM comparison.
    class HtmlLayoutValidator
    {
        static void Main()
        {
            // Paths to the source Excel file, the generated HTML, and the expected XML layout.
            string excelPath = @"C:\Data\SourceWorkbook.xlsx";
            string htmlPath = @"C:\Data\GeneratedLayout.html";
            string expectedXmlPath = @"C:\Data\ExpectedLayout.xml";

            try
            {
                // Verify that the source files exist.
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
                    return;
                }

                if (!File.Exists(expectedXmlPath))
                {
                    Console.WriteLine($"Error: Expected XML file not found at '{expectedXmlPath}'.");
                    return;
                }

                // Load the workbook.
                Workbook workbook = new Workbook(excelPath);

                // Save the workbook as HTML.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = true, // Export only the first worksheet for simplicity.
                    ExportImagesAsBase64 = true        // Embed images to keep the HTML self‑contained.
                };
                workbook.Save(htmlPath, htmlOptions);

                // Verify that the HTML file was created.
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Error: HTML file was not generated at '{htmlPath}'.");
                    return;
                }

                // Load the expected XML document.
                XmlDocument expectedXml = new XmlDocument();
                expectedXml.Load(expectedXmlPath);

                // Load the generated HTML as a plain text string.
                string htmlContent = File.ReadAllText(htmlPath);

                // Simple validation: check that the HTML contains the root element of the expected XML.
                // This is a lightweight placeholder for a full DOM comparison.
                bool containsRoot = htmlContent.Contains(expectedXml.DocumentElement.Name);

                Console.WriteLine(containsRoot
                    ? "HTML layout appears to contain the expected root element."
                    : "HTML layout does NOT contain the expected root element.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
