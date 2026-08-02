using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsXmlMapSummary
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook that contains XML maps
            // Replace "SourceWorkbook.xlsx" with the actual file path
            Workbook sourceWorkbook = new Workbook("SourceWorkbook.xlsx");

            // Create a new workbook to hold the summary report
            Workbook reportWorkbook = new Workbook();
            Worksheet reportSheet = reportWorkbook.Worksheets[0];

            // Write header row
            reportSheet.Cells["A1"].PutValue("XML Map Name");
            reportSheet.Cells["B1"].PutValue("Root Element");
            reportSheet.Cells["C1"].PutValue("Linked Cells Count");

            int reportRow = 1; // zero‑based index; start after header

            // Iterate through each XML map in the source workbook
            XmlMapCollection xmlMaps = sourceWorkbook.Worksheets.XmlMaps;
            for (int i = 0; i < xmlMaps.Count; i++)
            {
                XmlMap map = xmlMaps[i];
                string mapName = map.Name;
                string rootElement = map.RootElementName;
                int linkedCellCount = 0;

                // Scan all worksheets to find cells linked to this map
                foreach (Worksheet ws in sourceWorkbook.Worksheets)
                {
                    // Query cells linked to the root element path of the map
                    // The path format uses a leading slash
                    string queryPath = "/" + rootElement;
                    ArrayList cellAreas = ws.XmlMapQuery(queryPath, map);

                    // Each CellArea may cover multiple cells; count them all
                    foreach (CellArea area in cellAreas)
                    {
                        int rows = area.EndRow - area.StartRow + 1;
                        int cols = area.EndColumn - area.StartColumn + 1;
                        linkedCellCount += rows * cols;
                    }
                }

                // Write the map information into the report sheet
                reportSheet.Cells[reportRow, 0].PutValue(mapName);
                reportSheet.Cells[reportRow, 1].PutValue(rootElement);
                reportSheet.Cells[reportRow, 2].PutValue(linkedCellCount);
                reportRow++;
            }

            // Save the summary report workbook
            // Replace "XmlMapSummaryReport.xlsx" with the desired output path
            reportWorkbook.Save("XmlMapSummaryReport.xlsx");
        }
    }
}