// Title: Export a single worksheet from an .xlsx file to Excel 2003 XML while keeping data types and formatting using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook, copies the first worksheet into a new workbook, and saves it as Excel 2003 XML preserving data types and formatting with Aspose.Cells. | Show how to maintain the original worksheet name after copying and exporting to XML in Aspose.Cells. | Add robust error handling for missing input files and exceptions during the XML export process in C#.
// Common Searches: Aspose.Cells export first sheet to Excel 2003 XML preserving formatting | C# save worksheet as XML with original data types using Aspose.Cells | How to keep worksheet name when converting .xlsx to XML in .NET | Error handling for missing Excel file during Aspose.Cells XML export
// Tags: export worksheet to XML Aspose.Cells | preserve cell data types Aspose.Cells | retain original sheet name C# | copy worksheet to new workbook Aspose.Cells | SaveFormat.Xml usage .NET | Excel 2003 XML conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads 'input.xlsx', copies its first worksheet into a new workbook, retains the original sheet name, and saves the result as Excel 2003 XML ('output.xml') using Aspose.Cells, with file existence checks and basic exception handling.
class ExportWorksheetToXml
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xml";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook srcWorkbook = new Workbook(inputPath);
            Worksheet srcWorksheet = srcWorkbook.Worksheets[0]; // first worksheet

            // Create a new workbook and remove its default empty sheet
            Workbook destWorkbook = new Workbook();
            destWorkbook.Worksheets.Clear();

            // Copy the source worksheet into the destination workbook
            // AddCopy returns the index of the newly added worksheet
            int copiedIndex = destWorkbook.Worksheets.AddCopy(srcWorksheet.Name);
            Worksheet copiedWorksheet = destWorkbook.Worksheets[copiedIndex];
            // Preserve the original name (AddCopy may append a suffix)
            copiedWorksheet.Name = srcWorksheet.Name;

            // Save as Excel 2003 XML
            destWorkbook.Save(outputPath, SaveFormat.Xml);

            Console.WriteLine($"Worksheet exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
