using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetDocumentVersion
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Configure the document version property
            workbook.BuiltInDocumentProperties.Version = "1.0";

            // Display the configured version
            Console.WriteLine("Document Version: " + workbook.BuiltInDocumentProperties.Version);

            // Save the workbook as an XLSX file
            workbook.Save("DocumentVersionDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}