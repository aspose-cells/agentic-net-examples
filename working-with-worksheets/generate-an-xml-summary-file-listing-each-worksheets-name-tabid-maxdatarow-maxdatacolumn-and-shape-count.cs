// Title: Generate an XML summary of each worksheet’s name, TabId, MaxDataRow, MaxDataColumn, and shape count with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that iterates through all worksheets in a workbook and creates an XML document containing the worksheet Name, TabId, MaxDataRow, MaxDataColumn, and ShapeCount attributes. | Add comprehensive error handling to the worksheet summary generator to detect missing input files, insufficient permissions, and to log the full path of the generated XML file. | Extend the XML output schema to also record the number of charts and embedded pictures present on each worksheet.
// Common Searches: aspocells export worksheet properties to xml c# | how to get worksheet TabId and shape count with Aspose.Cells | retrieve max data row and column indices from Excel sheet using Aspose.Cells | xml report of all worksheets in a .NET console application | list worksheet metadata including shapes and charts using Aspose.Cells
// Tags: Aspose.Cells export worksheet metadata to XML | C# retrieve worksheet TabId with Aspose.Cells | Aspose.Cells get MaxDataRow and MaxDataColumn | count worksheet shapes Aspose.Cells | generate XML summary of Excel worksheets .NET | Aspose.Cells include chart count in worksheet report

using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;

// Creates an XML file that lists each worksheet's name, TabId, maximum data row/column indices, and shape count from an input Excel workbook using Aspose.Cells for .NET.
class WorksheetSummaryGenerator
{
    static void Main()
    {
        // Path to the Excel file to process
        string excelPath = "input.xlsx";

        // Path where the XML summary will be saved
        string xmlOutputPath = "summary.xml";

        // Load the workbook
        Workbook workbook = new Workbook(excelPath);

        // Create the root element for the XML document
        XElement root = new XElement("Worksheets");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Gather required information
            string name = sheet.Name;
            int tabId = sheet.TabId;
            int maxDataRow = sheet.Cells.MaxDataRow;       // Zero‑based index of the last row with data
            int maxDataColumn = sheet.Cells.MaxDataColumn; // Zero‑based index of the last column with data
            int shapeCount = sheet.Shapes.Count;

            // Build an XML element for the current worksheet
            XElement sheetElement = new XElement("Worksheet",
                new XAttribute("Name", name),
                new XAttribute("TabId", tabId),
                new XAttribute("MaxDataRow", maxDataRow),
                new XAttribute("MaxDataColumn", maxDataColumn),
                new XAttribute("ShapeCount", shapeCount)
            );

            // Add the worksheet element to the root
            root.Add(sheetElement);
        }

        // Create the XDocument and save it to the specified path
        XDocument doc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
        doc.Save(xmlOutputPath);

        Console.WriteLine($"Worksheet summary saved to '{Path.GetFullPath(xmlOutputPath)}'.");
    }
}
