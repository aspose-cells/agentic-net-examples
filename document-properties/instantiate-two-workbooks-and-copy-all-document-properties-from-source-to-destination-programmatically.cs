// Title: Copy all built‑in and custom document properties between Aspose.Cells workbooks (C#)
// Description: Demonstrates how to create a source workbook, set built‑in (Author, Title) and custom (Project, Version) properties, then programmatically copy every property to a new workbook and save the result using Aspose.Cells for .NET.
// Keywords: Aspose.Cells copy document properties | C# copy workbook metadata | transfer Excel custom properties | Aspose.Cells built‑in properties | copy Excel workbook properties programmatically | Aspose.Cells DocumentProperty example | .NET Excel metadata migration
// Common Searches: how to copy document properties with Aspose.Cells C# | copy built‑in and custom Excel properties programmatically | Aspose.Cells transfer workbook metadata | C# copy Excel file properties to another workbook | Aspose.Cells copy custom document properties
// Developer Intent: Programmatically duplicate every built‑in and custom document property from one Aspose.Cells workbook to another.
// Use Cases: Generate a fresh report from a template while preserving author, title, and project metadata. | Migrate custom metadata (e.g., Project, Version) from legacy Excel files to new analysis workbooks. | Synchronize document properties across multiple workbooks before archiving for compliance.
// AI Prompts: Provide a C# method that copies all built‑in and custom document properties from a source Aspose.Cells Workbook to a destination Workbook, handling missing properties gracefully. | Show code that iterates over source.BuiltInDocumentProperties and source.CustomDocumentProperties and replicates them in another workbook, converting non‑string values to strings when adding custom properties. | Explain how to update existing custom properties in the destination workbook while adding any that do not already exist.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Demonstrates how to create a source workbook, set built‑in (Author, Title) and custom (Project, Version) properties, then programmatically copy every property to a new workbook and save the result using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create the source workbook and set some document properties
            Workbook source = new Workbook();

            // Built‑in properties
            source.BuiltInDocumentProperties["Author"].Value = "John Smith";
            source.BuiltInDocumentProperties["Title"].Value = "Source Workbook";

            // Custom properties (convert non‑string values to string to match overload)
            source.CustomDocumentProperties.Add("Project", "AsposeDemo");
            source.CustomDocumentProperties.Add("Version", "1.0");

            // Save the source workbook (optional, just to have a file)
            string sourcePath = "Source.xlsx";
            source.Save(sourcePath);

            // Create an empty destination workbook
            Workbook destination = new Workbook();

            // ----- Copy Built‑in Document Properties -----
            foreach (DocumentProperty prop in source.BuiltInDocumentProperties)
            {
                // The property already exists in the destination collection; assign its value
                destination.BuiltInDocumentProperties[prop.Name].Value = prop.Value;
            }

            // ----- Copy Custom Document Properties -----
            foreach (DocumentProperty prop in source.CustomDocumentProperties)
            {
                // If the property already exists, update its value; otherwise add it
                if (destination.CustomDocumentProperties.Contains(prop.Name))
                {
                    destination.CustomDocumentProperties[prop.Name].Value = prop.Value;
                }
                else
                {
                    // Ensure the value is a string as required by the Add overload
                    destination.CustomDocumentProperties.Add(prop.Name, prop.Value?.ToString() ?? string.Empty);
                }
            }

            // Save the destination workbook with the copied properties
            string destinationPath = "Destination.xlsx";
            destination.Save(destinationPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
