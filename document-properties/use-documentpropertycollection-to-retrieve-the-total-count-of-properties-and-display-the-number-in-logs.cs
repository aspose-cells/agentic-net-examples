// Title: Count and Log Built‑in & Custom Document Properties with Aspose.Cells for .NET (C#)
// Description: Creates an empty workbook, reads the BuiltInDocumentProperties and CustomDocumentProperties collections, adds their counts, writes the combined total to the console, and optionally saves the file.
// Keywords: Aspose.Cells document property count | C# Aspose.Cells BuiltInDocumentProperties | Aspose.Cells CustomDocumentProperties | Workbook property total .NET | log document properties Aspose | Aspose.Cells API count properties | GitHub Aspose.Cells examples C#
// Common Searches: how to count built‑in document properties Aspose.Cells C# | Aspose.Cells get total custom properties | log workbook property count .NET | Aspose.Cells count properties example | C# retrieve document property collection size
// Developer Intent: Obtain the combined count of built‑in and custom workbook properties and output it.
// Use Cases: Verify that a generated spreadsheet contains the expected number of metadata entries before distribution. | Create an audit trail that records property counts for compliance or quality checks. | Trigger conditional logic (e.g., add or prune properties) based on the current total count.
// AI Prompts: Show code to enumerate each document property after calculating the total count with Aspose.Cells. | Generate a JSON log that includes the total property count plus name/value pairs for all properties. | Explain how to handle scenarios where the property count exceeds a predefined limit in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Creates an empty workbook, reads the BuiltInDocumentProperties and CustomDocumentProperties collections, adds their counts, writes the combined total to the console, and optionally saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (empty workbook)
        Workbook workbook = new Workbook();

        // Access the built‑in document properties collection
        DocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

        // Access the custom document properties collection
        DocumentPropertyCollection customProps = workbook.Worksheets.CustomDocumentProperties;

        // Retrieve the total number of properties (built‑in + custom)
        int totalPropertyCount = builtInProps.Count + customProps.Count;

        // Display the count in the console (log)
        Console.WriteLine($"Total document properties count: {totalPropertyCount}");

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("DocumentPropertiesCount.xlsx");
    }
}
