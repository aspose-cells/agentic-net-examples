// Title: Export a single worksheet from an XLSX workbook to an XML Spreadsheet with data types and formatting preserved using Aspose.Cells for .NET
// Common Searches: how to save a specific Excel worksheet as XML Spreadsheet with Aspose.Cells in C# | Aspose.Cells export worksheet to XML while keeping cell formatting and data types | C# copy a worksheet to a new workbook and save as .xml using Aspose.Cells | preserve number formats when converting Excel to XML Spreadsheet with Aspose.Cells | export first sheet of workbook to XML file using Aspose.Cells .NET API
// Tags: Aspose.Cells selective sheet XML export | XML Spreadsheet generation preserving cell data types | C# workbook worksheet copy for XML conversion | SaveFormat.Xml with formatting retention | Excel to XML conversion retaining number formats

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing XLSX file, extracts the first worksheet, copies it into a fresh workbook, removes the default sheet, and saves the result as an XML Spreadsheet using Aspose.Cells, which keeps the original cell data types and formatting.
class ExportWorksheetToXml
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xml";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(inputPath);

            // Select the worksheet to export (first worksheet in this example)
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Create a new workbook for the export
            Workbook exportWorkbook = new Workbook();

            // Remove the default empty worksheet
            exportWorkbook.Worksheets.Clear();

            // Copy the selected worksheet into the new workbook using the array overload
            exportWorkbook.Worksheets.AddCopy(
                new Worksheet[] { sourceSheet },
                new string[] { sourceSheet.Name });

            // Save the new workbook as an XML Spreadsheet
            exportWorkbook.Save(outputPath, SaveFormat.Xml);

            Console.WriteLine($"Worksheet exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
