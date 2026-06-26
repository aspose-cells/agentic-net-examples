using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsMetadataUtility
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (adjust as needed)
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "OutputWorkbook.xlsx";

            // Load the existing workbook (load rule)
            Workbook workbook = new Workbook(inputPath);

            // Access the custom document properties collection
            CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

            // Create (or replace) a worksheet named "Metadata"
            int metadataSheetIndex = workbook.Worksheets.Add();
            Worksheet metadataSheet = workbook.Worksheets[metadataSheetIndex];
            metadataSheet.Name = "Metadata";

            // Write header row
            Cells cells = metadataSheet.Cells;
            cells["A1"].PutValue("Property Name");
            cells["B1"].PutValue("Property Value");

            // Populate the worksheet with custom properties
            int row = 1; // zero‑based index; row 1 is the second row (after header)
            foreach (DocumentProperty prop in customProps)
            {
                cells[row, 0].PutValue(prop.Name);
                cells[row, 1].PutValue(prop.Value?.ToString() ?? string.Empty);
                row++;
            }

            // Save the modified workbook (save rule)
            workbook.Save(outputPath);
        }
    }
}