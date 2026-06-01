using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AddOleObjectWithHyperlinkDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Path to the source file that will be linked and opened via hyperlink
                string sourceFilePath = "sample.docx";

                // Ensure the source file exists before proceeding
                if (!File.Exists(sourceFilePath))
                {
                    Console.WriteLine($"Source file not found: {sourceFilePath}");
                    return;
                }

                // Read the file bytes to use as the OLE object's image (icon)
                byte[] imageData = File.ReadAllBytes(sourceFilePath);

                // Add a linked OLE object (topRow, leftColumn, height, width, imageData, linkedFile)
                int oleIndex = sheet.OleObjects.Add(5, 2, 200, 300, imageData, sourceFilePath);

                // Retrieve the added OleObject and mark it as a link
                OleObject ole = sheet.OleObjects[oleIndex];
                ole.IsLink = true;

                // Set the hyperlink address (Hyperlink property is read‑only; modify its Address)
                ole.Hyperlink.Address = sourceFilePath;

                // Save the workbook
                string outputPath = "OleObjectWithHyperlink.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}