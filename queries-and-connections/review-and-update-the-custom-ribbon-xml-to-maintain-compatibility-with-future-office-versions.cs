// Title: C# example to upgrade customUI Ribbon XML namespace and set OfficeVersion for future Office compatibility using Aspose.Cells
// Common Searches: how to upgrade customUI namespace in Ribbon.xml to the 2009/07 Office schema with C# | C# code to add or update OfficeVersion attribute in Ribbon XML for Office 2016 compatibility | using Aspose.Cells to load an Excel workbook while updating Ribbon XML in a C# application | replace root element namespace in customUI Ribbon XML programmatically in .NET | preserve whitespace when modifying Ribbon XML with XmlDocument in C#
// Tags: customUI namespace upgrade C# | OfficeVersion attribute set C# | Ribbon XML compatibility update | Aspose.Cells workbook loading example | XmlDocument preserve whitespace handling

using System;
using System.IO;
using System.Xml;
using Aspose.Cells;

namespace RibbonUpdaterExample
{
    // Demonstrates a C# utility that loads a Ribbon XML file with XmlDocument (preserving whitespace), verifies the <customUI> root, replaces its namespace with the latest Office 2009/07 schema, ensures the OfficeVersion attribute is set to 16.0 (Office 2016+), and saves the file. The sample also shows optional loading of an Excel workbook using Aspose.Cells.
    public static class RibbonUpdater
    {
        /// <param name="filePath">Full path to the Ribbon XML file.</param>
        public static void UpdateRibbonXml(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path must be provided.", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("Ribbon XML file not found.", filePath);

            try
            {
                // Load the existing Ribbon XML.
                var xmlDoc = new XmlDocument { PreserveWhitespace = true };
                xmlDoc.Load(filePath);

                // The root element should be <customUI>.
                XmlElement root = xmlDoc.DocumentElement;
                if (root == null || !root.Name.Equals("customUI", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Invalid Ribbon XML: Root element must be <customUI>.");

                // Update namespace to the latest Office namespace if needed.
                const string latestNamespace = "http://schemas.microsoft.com/office/2009/07/customui";

                if (root.NamespaceURI != latestNamespace)
                {
                    // Create a new root element with the desired namespace.
                    XmlElement newRoot = xmlDoc.CreateElement(root.Prefix, root.LocalName, latestNamespace);

                    // Copy existing attributes (excluding old xmlns attributes).
                    foreach (XmlAttribute attr in root.Attributes)
                    {
                        if (attr.Name.StartsWith("xmlns", StringComparison.OrdinalIgnoreCase))
                            continue;

                        XmlAttribute newAttr = (XmlAttribute)xmlDoc.ImportNode(attr, true);
                        newRoot.Attributes.Append(newAttr);
                    }

                    // Move all child nodes to the new root.
                    while (root.HasChildNodes)
                    {
                        XmlNode child = root.FirstChild;
                        root.RemoveChild(child);
                        newRoot.AppendChild(child);
                    }

                    // Replace the old root with the new one.
                    xmlDoc.ReplaceChild(newRoot, root);
                    root = newRoot;
                }

                // Ensure the OfficeVersion attribute is present and set to a high version.
                const string officeVersionAttr = "OfficeVersion";
                const string targetVersion = "16.0"; // Office 2016 and later.

                if (root.HasAttribute(officeVersionAttr))
                {
                    string existing = root.GetAttribute(officeVersionAttr);
                    if (Version.TryParse(existing, out Version existingVer) &&
                        Version.TryParse(targetVersion, out Version targetVer) &&
                        existingVer < targetVer)
                    {
                        root.SetAttribute(officeVersionAttr, targetVersion);
                    }
                }
                else
                {
                    root.SetAttribute(officeVersionAttr, targetVersion);
                }

                // Save the updated XML back to the file.
                xmlDoc.Save(filePath);
            }
            catch (Exception ex)
            {
                // Wrap exception to provide context.
                throw new ApplicationException($"Failed to update Ribbon XML file '{filePath}'.", ex);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Paths can be supplied via command line arguments or defaulted.
            string ribbonPath = args.Length > 0 ? args[0] : "Ribbon.xml";
            string workbookPath = args.Length > 1 ? args[1] : "Sample.xlsx";

            try
            {
                // Verify Ribbon XML file exists.
                if (!File.Exists(ribbonPath))
                {
                    Console.WriteLine($"Ribbon XML file not found: {ribbonPath}");
                    return;
                }

                // Demonstrate Aspose.Cells usage by loading a workbook (optional).
                if (File.Exists(workbookPath))
                {
                    Workbook wb = new Workbook(workbookPath);
                    Console.WriteLine($"Workbook loaded: {workbookPath} (Sheets: {wb.Worksheets.Count})");
                }
                else
                {
                    Console.WriteLine($"Workbook file not found (optional): {workbookPath}");
                }

                // Update the Ribbon XML.
                RibbonUpdater.UpdateRibbonXml(ribbonPath);
                Console.WriteLine("Ribbon XML updated successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
