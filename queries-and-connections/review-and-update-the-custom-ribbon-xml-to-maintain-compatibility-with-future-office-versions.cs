// Title: C# console tool to update Office custom ribbon XML for future version compatibility
// Common Searches: how to programmatically set OfficeVersion attribute in custom ribbon XML using C# | C# script to add missing xmlns namespace to Office ribbon XML file | update all button size attributes to large in Office ribbon XML with XDocument | automate compatibility changes in Office custom UI XML via console application | modify Office ribbon XML for newer Office 365 versions using C#
// Tags: update Office ribbon XML version attribute | add xmlns namespace to custom UI XML | set button size large in ribbon XML | C# XDocument modify Office ribbon file | console utility for Office ribbon compatibility

using System;
using System.IO;
using System.Xml.Linq;

namespace RibbonXmlUpdater
{
    // A C# console application that loads a custom ribbon XML file, ensures the root element contains an OfficeVersion attribute set to the latest known version (e.g., 16.0), adds a missing xmlns namespace declaration if needed, and forces every <button> element to have size="large". The tool saves the modified XML back to the original location and reports success or any processing errors.
    class Program
    {
        // Entry point of the console application
        static void Main(string[] args)
        {
            // Validate arguments: first argument is the path to the ribbon XML file
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: RibbonXmlUpdater <path-to-ribbon-xml>");
                return;
            }

            string ribbonPath = args[0];

            if (!File.Exists(ribbonPath))
            {
                Console.WriteLine($"Error: File not found - {ribbonPath}");
                return;
            }

            try
            {
                // Load the existing ribbon XML
                XDocument ribbonDoc = XDocument.Load(ribbonPath);

                // Namespace used by Office custom UI (may be present in the file)
                XNamespace ns = ribbonDoc.Root.GetDefaultNamespace();

                // Ensure the root element has the OfficeVersion attribute set to the latest known version.
                // This helps maintain compatibility with future Office releases.
                // If the attribute already exists, update its value; otherwise, add it.
                XAttribute versionAttr = ribbonDoc.Root.Attribute("OfficeVersion");
                if (versionAttr != null)
                {
                    versionAttr.Value = "16.0"; // Example: set to Office 2016/365 version
                }
                else
                {
                    ribbonDoc.Root.Add(new XAttribute("OfficeVersion", "16.0"));
                }

                // Optional: Add the "xmlns" attribute for future compatibility if missing.
                // Some older ribbon XML files may lack the explicit namespace declaration.
                XAttribute xmlnsAttr = ribbonDoc.Root.Attribute("xmlns");
                if (xmlnsAttr == null && !string.IsNullOrEmpty(ns.NamespaceName))
                {
                    ribbonDoc.Root.Add(new XAttribute("xmlns", ns.NamespaceName));
                }

                // Example: Ensure that all <button> elements have the "size" attribute set to "large"
                // This is a common requirement for newer Office UI guidelines.
                foreach (XElement button in ribbonDoc.Descendants(ns + "button"))
                {
                    XAttribute sizeAttr = button.Attribute("size");
                    if (sizeAttr == null)
                    {
                        button.Add(new XAttribute("size", "large"));
                    }
                    else
                    {
                        sizeAttr.Value = "large";
                    }
                }

                // Save the updated ribbon XML back to the original file (or you could choose a new file)
                ribbonDoc.Save(ribbonPath);
                Console.WriteLine("Ribbon XML updated successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the ribbon XML: {ex.Message}");
            }
        }
    }
}
