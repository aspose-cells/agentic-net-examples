// Title: How to embed a Visio diagram as a locked OLE object sized to specific cells using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# embed Visio .vsdx as OLE object and lock it in Excel | set OLE object dimensions to match an Excel cell range with Aspose.Cells .NET | prevent resizing of an embedded Visio OLE object in a generated workbook using Aspose.Cells | add OLE object from a byte array to worksheet cells A1:E5 with Aspose.Cells example | adjust row height and column width for OLE object sizing in Aspose.Cells C#
// Tags: visio ole embedding Aspose.Cells | set ole object size by cell range Aspose.Cells | lock ole object in generated workbook Aspose.Cells | add ole object from byte array C# Aspose.Cells | configure row height column width for ole sizing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsOleExample
{
    // The example loads a Visio .vsdx file into a byte array, creates a new Workbook, sets row heights and column widths to define the desired OLE object size, adds the Visio data as an OLE object covering cells A1 to E5, locks the object to prevent user resizing or moving, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the Visio file to embed
                string visioFilePath = @"C:\Path\To\YourDiagram.vsdx";

                // Verify that the Visio file exists
                if (!File.Exists(visioFilePath))
                    throw new FileNotFoundException("Visio file not found.", visioFilePath);

                // Load the Visio file into a byte array (required by OleObjects.Add)
                byte[] visioData = File.ReadAllBytes(visioFilePath);

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define the cell where the OLE object will be placed (e.g., A1)
                int startRow = 0;      // Row index (0‑based)
                int startColumn = 0;   // Column index (0‑based)

                // Set row heights and column widths to control OLE object size
                for (int i = startRow; i <= startRow + 4; i++)
                    sheet.Cells.SetRowHeight(i, 30); // height in points

                for (int j = startColumn; j <= startColumn + 4; j++)
                    sheet.Cells.SetColumnWidth(j, 15); // width in characters

                // Add the OLE object (Visio diagram) to the worksheet.
                // OleObjects.Add returns the index of the newly added object.
                int oleIndex = sheet.OleObjects.Add(
                    startRow,                 // upper‑left row
                    startColumn,              // upper‑left column
                    startRow + 4,             // lower‑right row (covers 5 rows)
                    startColumn + 4,          // lower‑right column (covers 5 columns)
                    visioData);               // OLE object data as byte[]

                // Retrieve the OleObject instance using the returned index
                OleObject ole = sheet.OleObjects[oleIndex];

                // Lock the OLE object so it cannot be moved or resized by the user
                ole.IsLocked = true;

                // Prepare output path and ensure directory exists
                string outputPath = @"C:\Path\To\Result.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
