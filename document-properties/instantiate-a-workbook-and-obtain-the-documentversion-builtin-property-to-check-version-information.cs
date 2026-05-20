using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Set the built‑in DocumentVersion property (optional)
        workbook.BuiltInDocumentProperties.DocumentVersion = "1.0";

        // Retrieve the DocumentVersion property value
        string docVersion = workbook.BuiltInDocumentProperties.DocumentVersion;

        // Output the version information
        Console.WriteLine("Document Version: " + docVersion);

        // Save the workbook to demonstrate the full lifecycle
        workbook.Save("DocumentVersionDemo.xlsx", SaveFormat.Xlsx);
    }
}