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
            string inputFile = "InputWorkbook.xlsx";
            string outputFile = "OutputWorkbook_WithMetadata.xlsx";

            // Load the existing workbook (lifecycle: load)
            Workbook workbook = new Workbook(inputFile);

            // Access the collection of custom document properties
            CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

            // Add a new worksheet named "Metadata"
            Worksheet metadataSheet = workbook.Worksheets.Add("Metadata");

            // Write header titles
            metadataSheet.Cells["A1"].PutValue("Property Name");
            metadataSheet.Cells["B1"].PutValue("Property Value");

            // Populate the worksheet with custom properties
            int rowIndex = 1; // zero‑based index; row 1 is the second row (after header)
            foreach (DocumentProperty prop in customProps)
            {
                metadataSheet.Cells[rowIndex, 0].PutValue(prop.Name);   // Column A
                metadataSheet.Cells[rowIndex, 1].PutValue(prop.Value); // Column B
                rowIndex++;
            }

            // Auto‑fit columns for better readability
            metadataSheet.AutoFitColumns();

            // Save the modified workbook (lifecycle: save)
            workbook.Save(outputFile);
        }
    }
}