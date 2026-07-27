using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Retrieve the built‑in DocumentVersion property
        string version = workbook.BuiltInDocumentProperties.DocumentVersion;

        // If the property is not set, assign a value for demonstration
        if (string.IsNullOrEmpty(version))
        {
            workbook.BuiltInDocumentProperties.DocumentVersion = "1.0";
            version = workbook.BuiltInDocumentProperties.DocumentVersion;
        }

        // Display the document version
        Console.WriteLine("Document Version: " + version);

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("DocumentVersionDemo.xlsx", SaveFormat.Xlsx);
    }
}