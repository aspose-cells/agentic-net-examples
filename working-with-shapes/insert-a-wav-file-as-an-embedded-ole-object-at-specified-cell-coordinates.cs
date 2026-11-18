// Title: Embedding a WAV audio file as an OLE object in a specific Excel cell with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells embed wav file as OLE object in Excel cell | How to add an audio OLE object to a worksheet at B2 using Aspose.Cells | Aspose.Cells C# insert OLE object from byte array and show as icon | Save Excel workbook with embedded WAV audio using Aspose.Cells | Set OleObject.IsLink false in Aspose.Cells C# example
// Tags: embed wav oleobject aspnet | add oleobject to worksheet cell c# | display oleobject as icon aspnet | oleobject islink false aspnet | save workbook with embedded audio aspnet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, reads a WAV file into a byte array, and inserts it as an embedded OLE object positioned at a specified cell (e.g., B2) with defined dimensions. The OLE object is set to be embedded (not linked) and displayed as an icon, then the workbook is saved as an XLSX file containing the embedded audio.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook with a default worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the WAV file that will be embedded
            string wavFilePath = @"C:\Temp\sound.wav";

            // Verify that the WAV file exists to avoid FileNotFoundException
            if (!File.Exists(wavFilePath))
            {
                Console.WriteLine($"Error: WAV file not found at '{wavFilePath}'.");
                return;
            }

            // Target cell where the OLE object will be placed (e.g., B2)
            // Row and column indices are zero‑based
            int targetRow = 1;      // Row 2
            int targetColumn = 1;   // Column B

            // Define the size of the OLE object (in pixels)
            int oleHeight = 100;
            int oleWidth = 200;

            // Read the WAV file into a byte array
            byte[] wavData;
            try
            {
                wavData = File.ReadAllBytes(wavFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read WAV file: {ex.Message}");
                return;
            }

            // Insert the WAV file as an embedded OLE object.
            // The overload requires left and top offsets (set to 0) and the binary data.
            OleObject ole = sheet.Shapes.AddOleObject(
                targetRow,
                targetColumn,
                oleHeight,
                oleWidth,
                0,          // left offset
                0,          // top offset
                wavData);

            // Ensure the object is embedded (not linked)
            ole.IsLink = false;

            // Display the OLE object as an icon
            ole.DisplayAsIcon = true;
            // Note: IconIndex property is not available in current Aspose.Cells version; default icon will be used.

            // Ensure output directory exists
            string outputPath = @"C:\Temp\EmbeddedAudio.xlsx";
            try
            {
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the embedded OLE object
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
