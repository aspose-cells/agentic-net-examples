using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DocumentVersionDemo
    {
        public static void Main()
        {
            // Create a new workbook (uses the provided Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Set the built‑in DocumentVersion property
            workbook.BuiltInDocumentProperties.DocumentVersion = "1.0";

            // Retrieve the DocumentVersion property
            string version = workbook.BuiltInDocumentProperties.DocumentVersion;

            // Output the version information
            Console.WriteLine("Document Version: " + version);

            // Save the workbook (uses the provided Save method rule)
            workbook.Save("DocumentVersionDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}