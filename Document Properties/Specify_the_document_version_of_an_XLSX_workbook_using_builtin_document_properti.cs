using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetDocumentVersion
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the document version built‑in property
            workbook.BuiltInDocumentProperties.DocumentVersion = "1.0";

            // Display the version (optional)
            Console.WriteLine("Document Version: " + workbook.BuiltInDocumentProperties.DocumentVersion);

            // Save the workbook as an XLSX file
            workbook.Save("DocumentVersionDemo.xlsx", SaveFormat.Xlsx);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            SetDocumentVersion.Run();
        }
    }
}