// Title: How to Get and Set the DocumentVersion Built‑In Property with Aspose.Cells for .NET (C#)
// Description: Demonstrates creating a Workbook with Aspose.Cells, reading the DocumentVersion built‑in property, updating it, and saving the file as XLSX. Useful for version tracking and audit of generated Excel spreadsheets in C# applications.
// Keywords: Aspose.Cells DocumentVersion | BuiltInDocumentProperties C# | read Excel document version | set Excel built‑in property | Aspose.Cells save workbook | C# Excel version control | Aspose.Cells .NET example
// Common Searches: Aspose.Cells get DocumentVersion property | set DocumentVersion built‑in property C# | read Excel file version with Aspose.Cells | how to change DocumentVersion in Aspose.Cells workbook | Aspose.Cells version tracking example
// Developer Intent: Read, modify, and persist the DocumentVersion built‑in property of an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Include a version number in generated reports for downstream consumers. | Maintain revision history by updating DocumentVersion after each automated edit. | Log the DocumentVersion value during batch processing for compliance audits.
// AI Prompts: Write C# code that opens an existing Excel file with Aspose.Cells, reads the DocumentVersion built‑in property, and prints it. | Show how to assign a custom DocumentVersion string to a Workbook and save it as an XLSX file using Aspose.Cells for .NET. | Explain a workflow for checking and updating the DocumentVersion property to implement version control in an Aspose.Cells automation script.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates creating a Workbook with Aspose.Cells, reading the DocumentVersion built‑in property, updating it, and saving the file as XLSX. Useful for version tracking and audit of generated Excel spreadsheets in C# applications.
    public class DocumentVersionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the DocumentVersion built‑in property
                string version = workbook.BuiltInDocumentProperties.DocumentVersion;
                Console.WriteLine("Document Version: " + version);

                // Set a version and save the workbook
                workbook.BuiltInDocumentProperties.DocumentVersion = "1.0";
                workbook.Save("DocumentVersionDemo.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DocumentVersionDemo.Run();
        }
    }
}
