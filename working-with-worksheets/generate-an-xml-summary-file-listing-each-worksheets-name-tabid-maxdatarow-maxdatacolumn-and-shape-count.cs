using System;
using System.Xml;
using Aspose.Cells;

namespace AsposeCellsSummaryDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Path where the XML summary will be saved
            string outputPath = "summary.xml";

            // Load the workbook (uses the provided Workbook constructor)
            Workbook workbook = new Workbook(inputPath);

            // Create an XmlWriter to generate the summary file
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "    "
            };

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

                    // Maximum data row index (zero‑based)
                    writer.WriteElementString("MaxDataRow", sheet.Cells.MaxDataRow.ToString());

                    // Maximum data column index (zero‑based)
                    writer.WriteElementString("MaxDataColumn", sheet.Cells.MaxDataColumn.ToString());

                    // Number of drawing shapes on the sheet
                    writer.WriteElementString("ShapeCount", sheet.Shapes.Count.ToString());

                    writer.WriteEndElement(); // </Worksheet>
                }

                writer.WriteEndElement(); // </WorksheetsSummary>
                writer.WriteEndDocument();
            }

            // The XML summary file is now saved at 'outputPath'
            Console.WriteLine($"Summary XML generated at: {outputPath}");
        }
    }
}