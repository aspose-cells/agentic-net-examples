using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsVersionDemo
{
    class Program
    {
        static void Main()
        {
            // Scenario 1: Set a custom document version using BuiltInDocumentProperties.DocumentVersion
            Workbook wb1 = new Workbook();
            wb1.BuiltInDocumentProperties.DocumentVersion = "2.5";
            wb1.Worksheets[0].Cells["A1"].PutValue("DocumentVersion = 2.5");
            wb1.Save("DocumentVersionDemo.xlsx", SaveFormat.Xlsx);

            // Scenario 2: Set the application version that created the file using BuiltInDocumentProperties.Version
            Workbook wb2 = new Workbook();
            // The format must be "00.0000"
            wb2.BuiltInDocumentProperties.Version = "12.0000";
            wb2.Worksheets[0].Cells["A1"].PutValue("Application Version = 12.0000");
            wb2.Save("ApplicationVersionDemo.xlsx", SaveFormat.Xlsx);

            // Scenario 3: Define the OOXML compliance level (specification version) via Workbook.Settings.Compliance
            Workbook wb3 = new Workbook();
            wb3.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;
            wb3.Worksheets[0].Cells["A1"].PutValue("OOXML Compliance = Strict");
            wb3.Save("OoxmlComplianceDemo.xlsx", SaveFormat.Xlsx);

            // Scenario 4: Load an existing workbook and update its DocumentVersion property
            Workbook wb4 = new Workbook("DocumentVersionDemo.xlsx");
            wb4.BuiltInDocumentProperties.DocumentVersion = "3.0";
            wb4.Worksheets[0].Cells["B1"].PutValue("Updated DocumentVersion = 3.0");
            wb4.Save("DocumentVersionUpdated.xlsx", SaveFormat.Xlsx);
        }
    }
}