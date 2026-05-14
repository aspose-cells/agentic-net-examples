using System;
using Aspose.Cells;

class DocumentVersionDemo
{
    static void Main()
    {
        // Create a new workbook (uses the provided constructor rule)
        Workbook workbook = new Workbook();

        // Set the DocumentVersion built‑in property (optional, shows how to assign)
        workbook.BuiltInDocumentProperties.DocumentVersion = "1.0";

        // Retrieve the DocumentVersion property
        string version = workbook.BuiltInDocumentProperties.DocumentVersion;

        // Output the version information
        Console.WriteLine("Document Version: " + version);

        // Save the workbook (uses the provided save rule)
        workbook.Save("DocumentVersionDemo.xlsx", SaveFormat.Xlsx);
    }
}