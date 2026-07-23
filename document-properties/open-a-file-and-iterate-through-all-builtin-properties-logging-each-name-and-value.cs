// Title: C# – List All Built‑In Document Properties of an Excel Workbook with Aspose.Cells
// Description: The sample opens an Excel file using Aspose.Cells, retrieves the BuiltInDocumentProperties collection, iterates through each property, and prints the property name and its value to the console.
// Keywords: Aspose.Cells read built‑in properties | C# list Excel document metadata | enumerate workbook built‑in properties | Aspose.Cells console output | Excel file metadata extraction .NET
// Common Searches: how to get built‑in document properties with Aspose.Cells | C# code to enumerate Excel workbook metadata | list all Excel built‑in properties using Aspose | Aspose.Cells example for reading document properties
// Developer Intent: Open an Excel workbook and display every built‑in document property (name and value).
// Use Cases: Validate author, company, or creation date before processing the file. | Create an audit trail of workbook metadata for compliance checks. | Debug unexpected property values when generating Excel reports.
// AI Prompts: Write C# code that exports the built‑in document properties to a CSV file with Aspose.Cells. | Show how to filter the property list to show only Author, Title, and Subject. | Provide error‑handling examples for null or missing property values during enumeration.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// The sample opens an Excel file using Aspose.Cells, retrieves the BuiltInDocumentProperties collection, iterates through each property, and prints the property name and its value to the console.
class Program
{
    static void Main()
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
            Console.WriteLine($"{prop.Name}: {prop.Value}");
        }
    }
}
