// Title: Copy Built‑In and Custom Document Properties Between Two Aspose.Cells Workbooks (C#)
// Description: Demonstrates how to create a source Workbook, set built‑in (Author, Title) and custom (ReviewedBy, Revision) properties, instantiate an empty destination Workbook, transfer all built‑in properties, add or update each custom property according to its runtime type, and save the result.
// Keywords: Aspose.Cells | C# | copy document properties | built‑in workbook metadata | custom Excel properties | transfer workbook metadata | property type handling int double bool string | clone Excel file | preserve Excel metadata | Aspose.Cells document properties example
// Common Searches: How to copy built‑in and custom document properties with Aspose.Cells for .NET | Aspose.Cells C# copy workbook metadata to another workbook | Transfer custom Excel properties between files using Aspose.Cells | Copy workbook properties programmatically C# Aspose | Preserve author and title when cloning an Excel workbook
// Developer Intent: Programmatically duplicate all built‑in and custom document properties from one Aspose.Cells Workbook to another.
// Use Cases: Generate a new report from a template while retaining original author, title, and other metadata. | Create version‑controlled copies of spreadsheets, preserving reviewer names and revision numbers. | Migrate a batch of Excel files to a new repository without losing any custom or built‑in properties.
// AI Prompts: Show a concise Aspose.Cells C# code snippet that copies both built‑in and custom document properties between Workbook objects, handling int, double, bool, and string values. | Explain how to detect existing custom properties in the destination workbook before adding them when copying metadata with Aspose.Cells. | Provide best practices for preserving Excel metadata during automated workbook cloning using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsPropertyCopyDemo
{
    // Demonstrates how to create a source Workbook, set built‑in (Author, Title) and custom (ReviewedBy, Revision) properties, instantiate an empty destination Workbook, transfer all built‑in properties, add or update each custom property according to its runtime type, and save the result.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create (or load) the source workbook
                Workbook sourceWorkbook = new Workbook(); // replace with new Workbook("source.xlsx") to load a file

                // Set some built‑in document properties in the source workbook
                sourceWorkbook.BuiltInDocumentProperties["Author"].Value = "John Smith";
                sourceWorkbook.BuiltInDocumentProperties["Title"].Value = "Sample Workbook";

                // Add custom document properties using appropriate overloads
                sourceWorkbook.CustomDocumentProperties.Add("ReviewedBy", "Jane Doe");
                sourceWorkbook.CustomDocumentProperties.Add("Revision", 3);

                // Create the destination workbook (empty)
                Workbook destWorkbook = new Workbook();

                // ----- Copy Built‑in Document Properties -----
                foreach (DocumentProperty srcProp in sourceWorkbook.BuiltInDocumentProperties)
                {
                    // Built‑in properties always exist in the destination workbook
                    destWorkbook.BuiltInDocumentProperties[srcProp.Name].Value = srcProp.Value;
                }

                // ----- Copy Custom Document Properties -----
                foreach (DocumentProperty srcProp in sourceWorkbook.CustomDocumentProperties)
                {
                    if (destWorkbook.CustomDocumentProperties.Contains(srcProp.Name))
                    {
                        // Update existing property
                        destWorkbook.CustomDocumentProperties[srcProp.Name].Value = srcProp.Value;
                    }
                    else
                    {
                        // Add new property using the overload that matches the value type
                        switch (srcProp.Value)
                        {
                            case int intVal:
                                destWorkbook.CustomDocumentProperties.Add(srcProp.Name, intVal);
                                break;
                            case double doubleVal:
                                destWorkbook.CustomDocumentProperties.Add(srcProp.Name, doubleVal);
                                break;
                            case bool boolVal:
                                destWorkbook.CustomDocumentProperties.Add(srcProp.Name, boolVal);
                                break;
                            default:
                                destWorkbook.CustomDocumentProperties.Add(srcProp.Name, srcProp.Value?.ToString() ?? string.Empty);
                                break;
                        }
                    }
                }

                // Save the destination workbook to verify the properties were copied
                string destPath = "DestinationWorkbook.xlsx";
                destWorkbook.Save(destPath);
                Console.WriteLine($"Destination workbook saved to '{destPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
