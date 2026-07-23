// Title: Set the Title Built‑in Document Property of an Excel Workbook using Aspose.Cells for .NET
// Description: Demonstrates how to create a new Workbook with Aspose.Cells, assign a descriptive project name to the Title built‑in document property, output the value for verification, and save the file as an XLSX document.
// Keywords: Aspose.Cells | C# | .NET | Excel workbook title property | built‑in document properties | set Title property | save workbook as xlsx | metadata in Excel files
// Common Searches: how to set title property in Excel using Aspose.Cells C# | Aspose.Cells example for workbook metadata | set built‑in document properties with Aspose.Cells .NET | create Excel file and add title metadata programmatically | Aspose.Cells save workbook with custom title
// Developer Intent: Add a custom Title built‑in document property to a newly created workbook and persist the change to an XLSX file.
// Use Cases: Embedding project identifiers in financial or analytical reports for easy retrieval in document management systems. | Automating template generation where each workbook carries a predefined title for downstream processing. | Preparing archival Excel files with searchable metadata to improve compliance and audit workflows.
// AI Prompts: Generate C# code that creates an Excel workbook with Aspose.Cells, sets the Title built‑in document property to a supplied string, and saves the file. | Show how to read, modify, and confirm the Title property of an existing Excel workbook using Aspose.Cells for .NET. | Explain how to set multiple built‑in document properties (Title, Author, Subject) in a single Aspose.Cells workbook creation script.

using System;
using Aspose.Cells;

// Demonstrates how to create a new Workbook with Aspose.Cells, assign a descriptive project name to the Title built‑in document property, output the value for verification, and save the file as an XLSX document.
class Program
{
    static void Main()
    {
        // Create a new workbook (uses the Workbook constructor rule)
        Workbook workbook = new Workbook();

        // Set the Title built‑in document property to a descriptive project name
        workbook.BuiltInDocumentProperties.Title = "Project Alpha – Financial Report 2024";

        // Output the set title for verification
        Console.WriteLine("Workbook Title: " + workbook.BuiltInDocumentProperties.Title);

        // Save the workbook (uses the Save method rule)
        workbook.Save("ProjectAlphaReport.xlsx", SaveFormat.Xlsx);
    }
}
