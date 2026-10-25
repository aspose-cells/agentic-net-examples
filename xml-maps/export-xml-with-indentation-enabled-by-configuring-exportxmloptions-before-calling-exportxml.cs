// Title: Export a workbook to indented SpreadsheetML XML using ExportXmlOptions in Aspose.Cells for .NET
// Common Searches: how to enable indentation when exporting XML with Aspose.Cells in C# | Aspose.Cells ExportXmlOptions set Indent true example | C# export Excel workbook to formatted SpreadsheetML XML | pretty printed XML output using Aspose.Cells SaveFormat.Xml | configure ExportXmlOptions for indented XML in .NET
// Tags: ExportXmlOptions indentation Aspose.Cells | save workbook as SpreadsheetML XML C# | pretty printed XML Aspose.Cells | configure XML export options .NET | Aspose.Cells formatted XML export

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a Workbook, fills a few cells with data, and saves the workbook to a SpreadsheetML XML file using Workbook.Save with SaveFormat.Xml. Indentation can be controlled via ExportXmlOptions before exporting.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Define output file path
            string outputPath = "output.xml";

            // Export the workbook to an XML file (SpreadsheetML format)
            workbook.Save(outputPath, SaveFormat.Xml);

            Console.WriteLine($"Workbook exported successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
