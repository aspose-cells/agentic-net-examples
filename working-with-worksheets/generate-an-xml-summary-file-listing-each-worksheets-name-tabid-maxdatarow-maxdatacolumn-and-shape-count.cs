// Title: Generate an XML summary of worksheet name, TabId, data limits, and shape count with Aspose.Cells for .NET
// Description: Loads a workbook with Aspose.Cells, iterates through every worksheet, and uses System.Xml.XmlWriter to create an indented XML file that records each sheet's Name, TabId, MaxDataRow, MaxDataColumn, and the number of drawing shapes. The workbook itself remains unchanged.
// Keywords: Aspose.Cells XML worksheet summary | C# export worksheet metadata | TabId shape count Aspose | MaxDataRow MaxDataColumn XML | generate workbook report .NET
// Common Searches: Aspose.Cells write worksheet details to XML C# | export sheet name TabId shape count as XML | list MaxDataRow and MaxDataColumn for each sheet | create workbook structure report using Aspose.Cells | C# generate XML summary of all worksheets
// Developer Intent: Produce an XML file that lists each worksheet’s name, internal TabId, highest occupied row and column, and total shape count.
// Use Cases: Document workbook layout for compliance audits | Feed sheet metadata into a cataloging or search index | Identify sheets with excessive shapes for cleanup | Track structural changes across workbook versions
// AI Prompts: Write C# code with Aspose.Cells that outputs an XML file containing Name, TabId, MaxDataRow, MaxDataColumn, and ShapeCount for every worksheet. | Extend the XML summary to include each worksheet’s visibility status (Visible, Hidden, VeryHidden). | Add comprehensive error handling for missing input files and failures when writing the summary XML, and log detailed messages.

using System;
using System.Xml;
using Aspose.Cells;

// Loads a workbook with Aspose.Cells, iterates through every worksheet, and uses System.Xml.XmlWriter to create an indented XML file that records each sheet's Name, TabId, MaxDataRow, MaxDataColumn, and the number of drawing shapes. The workbook itself remains unchanged.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create an XML writer for the summary file
        using (XmlWriter xmlWriter = XmlWriter.Create("summary.xml", new XmlWriterSettings { Indent = true }))
        {
            xmlWriter.WriteStartDocument();
            xmlWriter.WriteStartElement("WorksheetsSummary");

            // Iterate through each worksheet and write required information
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                xmlWriter.WriteStartElement("Worksheet");

                // Worksheet name
                xmlWriter.WriteElementString("Name", sheet.Name);

                // Internal tab identifier
                xmlWriter.WriteElementString("TabId", sheet.TabId.ToString());

                // Maximum row and column that contain data
                xmlWriter.WriteElementString("MaxDataRow", sheet.Cells.MaxDataRow.ToString());
                xmlWriter.WriteElementString("MaxDataColumn", sheet.Cells.MaxDataColumn.ToString());

                // Number of drawing shapes on the sheet
                xmlWriter.WriteElementString("ShapeCount", sheet.Shapes.Count.ToString());

                xmlWriter.WriteEndElement(); // </Worksheet>
            }

            xmlWriter.WriteEndElement(); // </WorksheetsSummary>
            xmlWriter.WriteEndDocument();
        }

        // No additional save operation needed for the workbook itself
    }
}
