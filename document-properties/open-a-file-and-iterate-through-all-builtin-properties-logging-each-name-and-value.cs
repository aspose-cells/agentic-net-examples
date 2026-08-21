// Title: C# – List All Built‑In Document Properties of an Excel Workbook with Aspose.Cells
// Description: Loads an Excel file using Aspose.Cells for .NET, accesses the BuiltInDocumentPropertyCollection, iterates through each property, safely handles null values, and writes the property name and value to the console.
// Keywords: Aspose.Cells C# | read built‑in document properties | Excel metadata extraction | list workbook properties | document property collection .NET | handle null property values
// Common Searches: Aspose.Cells get built‑in properties C# | list Excel file metadata with Aspose | iterate workbook document properties .NET | C# code to display Excel built‑in properties | how to read Excel workbook properties using Aspose
// Developer Intent: Open an Excel workbook and output every built‑in document property (name and value) to the console.
// Use Cases: Create an audit log of workbook metadata for compliance checks. | Export property information to a report or CSV for documentation. | Verify required metadata such as Author, Title, or Company before processing the file.
// AI Prompts: Show how to write the built‑in properties to a CSV file instead of the console. | Provide code that skips properties with null values and logs only populated entries. | Explain how to combine custom and built‑in document properties in a single enumeration using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsBuiltInPropertiesDemo
{
    // Loads an Excel file using Aspose.Cells for .NET, accesses the BuiltInDocumentPropertyCollection, iterates through each property, safely handles null values, and writes the property name and value to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be opened
            string filePath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Retrieve the collection of built‑in document properties
            BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

            // Iterate through each property and log its name and value
            foreach (DocumentProperty prop in builtInProps)
            {
                // Some built‑in properties may have null values; handle gracefully
                string value = prop.Value != null ? prop.Value.ToString() : "null";
                Console.WriteLine($"{prop.Name}: {value}");
            }
        }
    }
}
