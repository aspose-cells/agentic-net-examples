using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetDocumentVersionDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the document version property
            workbook.BuiltInDocumentProperties.DocumentVersion = "1.0";

            // Display the document version
            Console.WriteLine("Document Version: " + workbook.BuiltInDocumentProperties.DocumentVersion);

            // Save the workbook as an XLSX file
            workbook.Save("DocumentVersionDemo.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetDocumentVersionDemo.Run();
        }
    }
}