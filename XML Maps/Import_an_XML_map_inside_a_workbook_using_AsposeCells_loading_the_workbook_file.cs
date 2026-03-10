using System;
using Aspose.Cells;

class ImportXmlMapExample
{
    static void Main()
    {
        // Paths to the existing workbook and the XML file to be imported
        string workbookPath = "input.xlsx";
        string xmlPath = "data.xml";

        // Destination sheet name and start cell (row 0, column 0 corresponds to A1)
        string sheetName = "Sheet1";
        int startRow = 0;
        int startCol = 0;

        // Load the existing workbook (lifecycle rule: workbook-load)
        Workbook workbook = new Workbook(workbookPath);

        // Import the XML data into the specified worksheet at the given cell (ImportXml method)
        workbook.ImportXml(xmlPath, sheetName, startRow, startCol);

        // Save the workbook with the imported XML map (lifecycle rule: workbook-save)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}