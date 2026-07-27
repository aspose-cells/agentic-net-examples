using System;
using System.Xml.Linq;
using Aspose.Cells;

namespace AsposeCellsNamedRangeExport
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Create the root element for the XML document
            XDocument xmlDoc = new XDocument(new XElement("NamedRanges"));

            // Iterate through all defined names in the workbook
            foreach (Name definedName in workbook.Worksheets.Names)
            {
                // Create an element for the current name
                XElement nameElement = new XElement("Name",
                    new XAttribute("Text", definedName.Text));

                // Retrieve all referred areas (including external links)
                ReferredArea[] areas = definedName.GetReferredAreas(true);

                if (areas != null)
                {
                    foreach (ReferredArea area in areas)
                    {
                        XElement areaElement = new XElement("Area",
                            new XAttribute("IsExternalLink", area.IsExternalLink),
                            new XAttribute("SheetName", area.SheetName ?? string.Empty),
                            new XAttribute("IsArea", area.IsArea));

                        // Add external file name if the area refers to an external workbook
                        if (area.IsExternalLink && !string.IsNullOrEmpty(area.ExternalFileName))
                        {
                            areaElement.Add(new XAttribute("ExternalFileName", area.ExternalFileName));
                        }

                        // Add cell or range information
                        if (area.IsArea)
                        {
                            string startCell = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                            string endCell = CellsHelper.CellIndexToName(area.EndRow, area.EndColumn);
                            areaElement.Add(new XAttribute("StartCell", startCell));
                            areaElement.Add(new XAttribute("EndCell", endCell));
                        }
                        else
                        {
                            string cell = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                            areaElement.Add(new XAttribute("Cell", cell));
                        }

                        nameElement.Add(areaElement);
                    }
                }

                xmlDoc.Root.Add(nameElement);
            }

            // Save the generated XML to a file
            xmlDoc.Save("NamedRanges.xml");

            Console.WriteLine("Named ranges exported to NamedRanges.xml");
        }
    }
}