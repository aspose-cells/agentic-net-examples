// Title: C# – Import XML File into an Aspose.Cells Workbook Using Workbook.ImportXml
// Description: Demonstrates how to create a new Workbook, import an XML file into a specified worksheet and cell (zero‑based indices), automatically create the sheet if it does not exist, and save the result as an XLSX file with Aspose.Cells.
// Keywords: Aspose.Cells ImportXml C# | import XML into Excel workbook | Workbook.ImportXml example | load XML data Aspose.Cells | create worksheet from XML | Aspose.Cells XML map import
// Common Searches: Aspose.Cells import XML file C# | Workbook.ImportXml start row column | how to add XML data to Excel with Aspose | create worksheet and import XML Aspose.Cells | C# code for ImportXml method
// Developer Intent: Load an XML file into a new or existing worksheet at a defined cell location using Aspose.Cells.
// Use Cases: Import a configuration XML into Sheet1 at A1 and generate a ready‑to‑share XLSX report. | Load product‑catalog XML files into separate worksheets for consolidated analysis. | Automate merging of multiple XML data sources into one workbook, each on its own sheet.
// AI Prompts: Generate C# code that uses Workbook.ImportXml to load an XML file into a specific worksheet and cell, then saves the workbook. | Show how to verify a worksheet exists before calling ImportXml and create it if missing. | Explain exception handling for invalid XML paths or malformed XML when using ImportXml.

using System;
using Aspose.Cells;

namespace AsposeCellsXmlImportDemo
{
    // Demonstrates how to create a new Workbook, import an XML file into a specified worksheet and cell (zero‑based indices), automatically create the sheet if it does not exist, and save the result as an XLSX file with Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Path to the XML file that contains the data to be imported
            string xmlFilePath = "data.xml";

            // Destination worksheet name (will be created if it does not exist)
            string sheetName = "Sheet1";

            // Starting cell for the import (row and column are zero‑based)
            int startRow = 0; // corresponds to Excel row 1 (A)
            int startColumn = 0; // corresponds to Excel column A

            // Import the XML data into the specified worksheet and cell
            workbook.ImportXml(xmlFilePath, sheetName, startRow, startColumn);

            // Save the workbook to an Excel file
            workbook.Save("ImportedXmlWorkbook.xlsx");
        }
    }
}
