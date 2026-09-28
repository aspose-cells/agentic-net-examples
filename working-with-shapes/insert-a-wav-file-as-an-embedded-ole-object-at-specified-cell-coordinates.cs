// Title: Embed a WAV audio file as an OLE object in a specific Excel cell using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells' Worksheet.OleObjects.Add to insert a WAV file from a byte array into cell B2 with custom height and width in C#. | Create a new workbook, read a .wav file into memory, embed it as an OLE object at row 1, column 1, and save the workbook as XLSX using Aspose.Cells.
// Common Searches: aspnet embed wav file into excel worksheet as ole object using aspose.cells | c# add audio ole object to specific cell in xlsx with aspose | how to set height and width of embedded audio object in excel via aspose.cells | reading wav file into byte array for oleobject insertion in asp.net | save workbook with embedded audio using aspose.cells c#
// Tags: embed wav ole object aspose.cells c# | worksheet oleobjects.add audio file | specify cell coordinates for ole object insertion | custom ole object dimensions pixels | handle wav file not found exception aspose | save workbook with embedded audio xlsx

using Aspose.Cells;
using System;
using System.IO;

// The example reads a WAV file into a byte array, creates a new workbook, and uses Worksheet.OleObjects.Add to embed the audio as an OLE object at row 1, column 1 (cell B2) with a size of 200 × 100 pixels, then saves the result as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the WAV file that will be embedded
            string wavFilePath = "sample.wav";

            // Verify that the WAV file exists to avoid FileNotFoundException
            if (!File.Exists(wavFilePath))
                throw new FileNotFoundException($"The specified WAV file was not found: {wavFilePath}");

            // Read the WAV file into a byte array (required by OleObjects.Add)
            byte[] wavData = File.ReadAllBytes(wavFilePath);

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Target cell coordinates (zero‑based). Example: B2 -> row 1, column 1
            int targetRow = 1;      // Row index
            int targetColumn = 1;   // Column index

            // Size of the OLE object in pixels (adjust as required)
            int oleHeight = 100;
            int oleWidth = 200;

            // Insert the WAV file as an embedded OLE object at the specified cell.
            sheet.OleObjects.Add(
                targetRow,      // Upper‑left row
                targetColumn,   // Upper‑left column
                oleHeight,      // Height in pixels
                oleWidth,       // Width in pixels
                wavData         // Byte array of the WAV file
            );

            // Save the workbook to a file
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
