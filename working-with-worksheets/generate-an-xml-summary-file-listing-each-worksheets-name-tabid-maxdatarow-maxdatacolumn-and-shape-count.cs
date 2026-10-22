// Title: Generate an XML file summarizing each worksheet’s name, TabId, MaxDataRow, MaxDataColumn, and shape count with Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# console application that loads an Excel workbook using Aspose.Cells, iterates through all worksheets, and creates an XML document containing the worksheet name, TabId, MaxDataRow, MaxDataColumn, and the number of shapes, then saves it to a specified file. | Using Aspose.Cells for .NET, produce code that extracts worksheet metadata (name, tab identifier, maximum data row/column indices, and shape count) and writes the information into a structured XML summary file. | Create a C# script that reads "input.xlsx", gathers each sheet’s properties via Cells.MaxDataRow, Cells.MaxDataColumn, and Shapes.Count, and outputs an XML file "summary.xml" with a <Worksheets> root element.
// Common Searches: Aspose.Cells C# export worksheet name, tab id, max data row and column to XML | How to retrieve shape count for each worksheet using Aspose.Cells in .NET | Create XML summary of all worksheets in an Excel workbook with Aspose.Cells | C# code to list worksheet properties and save them as an XML file | Aspose.Cells get MaxDataRow and MaxDataColumn for every sheet in a workbook
// Tags: Aspose.Cells export worksheet metadata to XML | C# retrieve worksheet MaxDataRow MaxDataColumn Aspose.Cells | Aspose.Cells count shapes per worksheet | generate Excel worksheet summary .NET | save worksheet properties as XML using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Xml.Linq;

// A C# console program that loads an Excel file with Aspose.Cells, iterates through each worksheet, captures the sheet name, TabId, maximum data row and column indices, and the count of shapes, then writes this information into a structured XML summary file.
class WorksheetSummaryGenerator
{
    static void Main()
    {
        // Path to the Excel file to process
        string excelPath = "input.xlsx";

        // Path where the XML summary will be saved
        string xmlPath = "summary.xml";

        // Verify that the input Excel file exists
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: The file '{excelPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Create the root element for the XML document
            XElement root = new XElement("Worksheets");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Use Cells property to get max data row/column
                int maxDataRow = sheet.Cells.MaxDataRow;
                int maxDataColumn = sheet.Cells.MaxDataColumn;

                // Build an XML element containing the required worksheet information
                XElement sheetElement = new XElement("Worksheet",
                    new XElement("Name", sheet.Name),
                    new XElement("TabId", sheet.TabId),
                    new XElement("MaxDataRow", maxDataRow),
                    new XElement("MaxDataColumn", maxDataColumn),
                    new XElement("ShapeCount", sheet.Shapes.Count)
                );

                // Add the worksheet element to the root
                root.Add(sheetElement);
            }

            // Create the XML document with a declaration and the root element
            XDocument doc = new XDocument(
                new XDeclaration("1.0", "utf-8", "yes"),
                root
            );

            // Save the XML summary to the specified file
            doc.Save(xmlPath);

            Console.WriteLine($"Worksheet summary saved to '{xmlPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
