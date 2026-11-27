// Title: Create an XML file that lists each worksheet’s name, TabId, maximum data row/column and shape count using Aspose.Cells for .NET (C#)
// Common Searches: C# code to export worksheet metadata (name, TabId, max row, max column, shape count) to XML with Aspose.Cells | How to generate an XML summary of all sheets in an Excel workbook using Aspose.Cells .NET | Aspose.Cells example for retrieving worksheet TabId and shape count and writing to XML | Write worksheet information to an XML file in C# with Aspose.Cells Workbook | Export Excel sheet properties (max data row/column) to XML using Aspose.Cells API
// Tags: Aspose.Cells worksheet metadata XML export | Aspose.Cells get worksheet TabId | Aspose.Cells max data row column | Aspose.Cells shape count per worksheet | C# generate worksheet summary XML

using System;
using System.IO;
using System.Xml;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, iterates through each worksheet, and writes an XML document containing the sheet's Name, internal TabId, maximum data row and column indices, and the total number of shapes on the sheet.
class Program
{
    static void Main(string[] args)
    {
        // Expect two arguments: input workbook path and output XML path
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: Program.exe <inputWorkbookPath> <outputXmlPath>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        // Validate input file existence
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' does not exist.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Prepare XML writer settings for pretty output
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                Encoding = System.Text.Encoding.UTF8
            };

            // Create the XML summary file
            using (XmlWriter writer = XmlWriter.Create(outputPath, settings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("WorksheetsSummary");

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    writer.WriteStartElement("Worksheet");

                    // Worksheet name
                    writer.WriteElementString("Name", sheet.Name);

                    // Internal TabId
                    writer.WriteElementString("TabId", sheet.TabId.ToString());

                    // Maximum data row and column (use Cells collection)
                    int maxRow = sheet.Cells.MaxDataRow;      // Returns -1 if no data
                    int maxCol = sheet.Cells.MaxDataColumn;   // Returns -1 if no data
                    writer.WriteElementString("MaxDataRow", maxRow.ToString());
                    writer.WriteElementString("MaxDataColumn", maxCol.ToString());

                    // Number of shapes on the worksheet
                    int shapeCount = sheet.Shapes.Count;
                    writer.WriteElementString("ShapeCount", shapeCount.ToString());

                    writer.WriteEndElement(); // </Worksheet>
                }

                writer.WriteEndElement(); // </WorksheetsSummary>
                writer.WriteEndDocument();
            }

            Console.WriteLine($"Summary XML successfully written to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Report any unexpected errors without suppressing them
            Console.WriteLine("An error occurred while processing the workbook:");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
}
