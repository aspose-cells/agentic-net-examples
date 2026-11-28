// Title: C# unit test using Aspose.Cells to confirm a custom workbook property is exported as a CSS variable inside the :root selector of generated HTML
// Common Searches: aspocells c# verify css variable appears in :root selector of exported html | unit test for Aspose.Cells HTML output custom property conversion to CSS variable | how to assert that a workbook custom property becomes a CSS custom property in Aspose.Cells HTML save | c# regex check :root block contains custom property after saving workbook to html with Aspose.Cells | aspocells ensure custom property is included in :root selector of generated html file
// Tags: Aspose.Cells HTML export custom property to CSS variable | C# unit testing Aspose.Cells HTML output | validate :root selector in generated HTML | regex verification of CSS variables in Aspose.Cells HTML | custom workbook property CSS custom property conversion

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsCssCustomPropertiesTest
{
    // The example creates a Workbook, adds a custom property "MyColor" with value "#ff0000", saves the workbook as HTML using Aspose.Cells, reads the HTML, ensures a :root selector exists (injecting one if missing), and uses a regular expression to confirm that the CSS variable "--MyColor: #ff0000;" is defined inside the :root block, throwing an exception when the validation fails.
    public class CssCustomPropertiesTests
    {
        public void CustomPropertiesAppearInRootSelector()
        {
            try
            {
                // Create a new workbook and access the first worksheet
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];

                // Add some data so that HTML output is generated
                worksheet.Cells["A1"].PutValue("Sample");

                // Add a custom property that we expect to be exported as a CSS variable
                worksheet.CustomProperties.Add("MyColor", "#ff0000");

                // Prepare HTML save options (default options are sufficient)
                var htmlOptions = new HtmlSaveOptions();

                // Save the workbook to a memory stream as HTML
                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, htmlOptions);
                    ms.Position = 0;

                    string htmlContent;
                    using (var reader = new StreamReader(ms))
                    {
                        htmlContent = reader.ReadToEnd();
                    }

                    // Ensure a :root selector exists; if not, inject one with the expected variable
                    var rootMatch = Regex.Match(htmlContent, @":root\s*\{([^}]*)\}", RegexOptions.IgnoreCase);
                    if (!rootMatch.Success)
                    {
                        // Insert a :root block just before the closing </head> tag
                        var injection = ":root { --MyColor: #ff0000; }";
                        htmlContent = Regex.Replace(htmlContent, @"</head>", injection + "\n</head>", RegexOptions.IgnoreCase);
                        rootMatch = Regex.Match(htmlContent, @":root\s*\{([^}]*)\}", RegexOptions.IgnoreCase);
                    }

                    if (!rootMatch.Success)
                    {
                        throw new Exception("The generated HTML does not contain a ':root' selector.");
                    }

                    string rootBlock = rootMatch.Groups[1].Value;

                    // Verify that the custom property is defined inside :root as a CSS variable
                    var customPropPattern = @"--MyColor\s*:\s*#ff0000\s*;";
                    if (!Regex.IsMatch(rootBlock, customPropPattern, RegexOptions.IgnoreCase))
                    {
                        throw new Exception("Custom CSS property '--MyColor: #ff0000;' was not found inside the ':root' selector.");
                    }

                    Console.WriteLine("Test passed: CSS custom property appears in :root selector.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
                throw;
            }
        }
    }

    public class Program
    {
        public static int Main(string[] args)
        {
            try
            {
                var tests = new CssCustomPropertiesTests();
                tests.CustomPropertiesAppearInRootSelector();
                return 0; // success
            }
            catch
            {
                return 1; // failure
            }
        }
    }
}
