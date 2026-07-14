using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // (Optional) Set the DocumentVersion property to a known value
        workbook.BuiltInDocumentProperties.DocumentVersion = "1.0";

        // Retrieve the DocumentVersion built‑in property
        string docVersion = workbook.BuiltInDocumentProperties.DocumentVersion;

        // Output the version information
        Console.WriteLine("Document Version: " + docVersion);

        // Save the workbook (required to persist any changes)
        workbook.Save("DocumentVersionDemo.xlsx", SaveFormat.Xlsx);
    }
}