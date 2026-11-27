// Title: Add a custom effect scheme to an Excel workbook’s theme using Aspose.Cells for .NET (C#)
// Common Searches: how to add a custom effect scheme to an Excel theme with Aspose.Cells in C# | extract and edit workbook theme XML using Aspose.Cells C# | apply modified theme XML to a workbook with Aspose.Cells | C# add innerShadow effect to Excel theme via Aspose.Cells | Aspose.Cells ThemeXml property example for custom effects
// Tags: Aspose.Cells add effectScheme to theme XML | C# modify Excel theme effect scheme | Aspose.Cells ThemeXml property usage | innerShadow element in Excel theme XML | update workbook theme with custom effect scheme

using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;

// The example loads an existing XLSX file, retrieves its theme XML through the Workbook.Theme dynamic object, parses the XML, creates a custom <effectScheme> containing an <innerShadow> definition, removes any existing scheme with the same name, appends the new scheme to the <themeElements> node, writes the modified XML back to the workbook via the ThemeXml property, and saves the updated workbook.
class ThemeEnhancer
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the workbook's theme using dynamic to avoid compile‑time dependency on Theme type
            dynamic theme = workbook.Theme;
            if (theme == null)
            {
                Console.WriteLine("The workbook does not contain a theme.");
                return;
            }

            string themeXml = theme.ThemeXml as string;
            if (string.IsNullOrEmpty(themeXml))
            {
                Console.WriteLine("The workbook's theme XML is empty.");
                return;
            }

            // Parse the theme XML
            XDocument themeDoc = XDocument.Parse(themeXml);
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

            // Locate the <a:themeElements> element
            XElement themeElements = themeDoc.Root.Element(a + "themeElements");
            if (themeElements == null)
            {
                throw new InvalidOperationException("The theme XML does not contain a themeElements node.");
            }

            // Create a custom effect scheme
            XElement customEffectScheme = new XElement(a + "effectScheme",
                new XAttribute("name", "CustomEffectScheme"),
                new XElement(a + "innerShadow",
                    new XAttribute("blurRad", "40000"),
                    new XAttribute("dist", "20000"),
                    new XAttribute("dir", "5400000"),
                    new XAttribute("sx", "100000"),
                    new XAttribute("sy", "100000"),
                    new XAttribute("kx", "0"),
                    new XAttribute("ky", "0"),
                    new XAttribute("algn", "tl"),
                    new XElement(a + "srgbClr", new XAttribute("val", "000000"))
                )
            );

            // Remove any existing scheme with the same name
            foreach (XElement scheme in themeElements.Elements(a + "effectScheme"))
            {
                if ((string)scheme.Attribute("name") == "CustomEffectScheme")
                {
                    scheme.Remove();
                    break;
                }
            }

            // Append the new scheme
            themeElements.Add(customEffectScheme);

            // Serialize the modified XML (preserve XML declaration if present)
            string modifiedThemeXml = themeDoc.Declaration != null
                ? themeDoc.Declaration.ToString() + themeDoc.ToString()
                : themeDoc.ToString();

            // Apply the modified theme back to the workbook
            theme.ThemeXml = modifiedThemeXml;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
