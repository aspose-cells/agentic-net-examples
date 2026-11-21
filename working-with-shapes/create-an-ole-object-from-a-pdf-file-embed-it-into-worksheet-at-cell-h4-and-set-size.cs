// Title: Embed a PDF file as an OLE object in cell H4 with custom height and width using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# embed PDF as OLE object at cell H4 | Set custom size for PDF OLE object in Excel with Aspose.Cells | How to add a PDF OLE object to a worksheet at a specific location in .NET | OleObjects.Add PDF byte array specific cell and dimensions Aspose.Cells | Create and position a PDF OLE object in Excel using C#
// Tags: embed pdf oleobject aspose.cells c# | oleobjects.add byte array | set oleobject dimensions worksheet | place oleobject at specific cell excel | pdf embedding using aspose.cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example demonstrates how to load a PDF file into a byte array and embed it as an OLE object in an Excel worksheet at cell H4 using Aspose.Cells for .NET. It sets the object's height to 200 points and width to 300 points, then saves the workbook as OutputWithPdfOle.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Initialize a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Path to the PDF file that will be embedded as an OLE object
                string pdfPath = "sample.pdf";

                // Ensure the PDF file exists before attempting to embed it
                if (!File.Exists(pdfPath))
                {
                    Console.WriteLine($"File not found: {pdfPath}");
                    return;
                }

                // Read the PDF file into a byte array (required by the OleObjects.Add overload)
                byte[] pdfData = File.ReadAllBytes(pdfPath);

                // Define the target cell (H4) – zero‑based indices: row 3, column 7
                int targetRow = 3;      // Row 4 in Excel (zero‑based)
                int targetColumn = 7;   // Column H in Excel (zero‑based)

                // Desired size of the OLE object (in points)
                int oleHeight = 200;    // Height of the embedded object
                int oleWidth = 300;     // Width of the embedded object

                // Add the OLE object to the worksheet at the specified location and size
                // Using overload that accepts a byte[] containing the file data
                sheet.OleObjects.Add(targetRow, targetColumn, oleHeight, oleWidth, pdfData);

                // Save the workbook to a file
                string outputPath = "OutputWithPdfOle.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
