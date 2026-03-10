using System;
using Aspose.Cells;

namespace AsposeCellsXmlImportDemo
{
    class Program
    {
        static void Main()
        {
            // Paths for the existing workbook and the XML file to import
            string workbookPath = "input.xlsx";
            string xmlPath = "data.xml";

            // Destination sheet name and start cell (row 0, column 0 corresponds to A1)
            string sheetName = "Sheet1";
            int startRow = 0;
            int startColumn = 0;

            // Load the existing workbook (lifecycle rule: workbook-load)
            Workbook workbook = new Workbook(workbookPath);

            // Import the XML data into the specified worksheet and cell (ImportXml method)
            workbook.ImportXml(xmlPath, sheetName, startRow, startColumn);

            // Save the workbook with the imported XML data (lifecycle rule: workbook-save)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}