// Title: How to embed a WAV audio file as an OLE object in a specific Excel cell using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# embed wav file into Excel cell as OLE object | Add audio OLE object to worksheet at B2 with Aspose.Cells | Set custom height and width for embedded WAV OLE object in Aspose.Cells | Insert OLE object from byte array into Excel using Shapes.AddOleObject | Save workbook after embedding audio file with Aspose.Cells C#
// Tags: insert audio ole object Aspose.Cells | add ole object to worksheet cell | Shapes.AddOleObject byte array usage | custom ole object dimensions .NET | load wav bytes for ole insertion

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, reads a WAV file into a byte array, and uses Worksheet.Shapes.AddOleObject to insert the audio as an OLE object at a specified cell with custom height and width, then saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the WAV file to embed
            string wavFilePath = @"C:\Path\To\Audio.wav";

            // Verify the WAV file exists
            if (!File.Exists(wavFilePath))
            {
                Console.WriteLine($"Audio file not found: {wavFilePath}");
                return;
            }

            // Target cell coordinates (e.g., B2)
            int targetRow = 1;      // Row index (0‑based) for row 2
            int targetColumn = 1;   // Column index (0‑based) for column B

            // Size of the OLE object (in pixels)
            int oleHeight = 100;
            int oleWidth = 200;

            // Read the audio file into a byte array
            byte[] oleData = File.ReadAllBytes(wavFilePath);

            // Insert the WAV file as an embedded OLE object at the specified cell
            // Offsets (0,0) place the object at the top‑left corner of the cell.
            sheet.Shapes.AddOleObject(targetRow, targetColumn, oleHeight, oleWidth, 0, 0, oleData);

            // Ensure output directory exists
            string outputPath = @"C:\Path\To\Output.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
