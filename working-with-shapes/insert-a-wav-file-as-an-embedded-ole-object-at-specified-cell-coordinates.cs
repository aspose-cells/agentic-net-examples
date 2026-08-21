// Title: C# – Embed a WAV file as an OLE object in a specific Excel cell with Aspose.Cells
// Description: Demonstrates how to read a WAV file (and optional PNG icon), add an OleObject at given row/column coordinates, embed the audio data with SetEmbeddedObject, set a label, and save the workbook as XLSX using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | embed WAV | OLE object | Excel audio | SetEmbeddedObject | custom icon | worksheet OleObject | source code example | GitHub | coding agent
// Common Searches: embed wav file in Excel using Aspose.Cells C# | add audio OLE object to worksheet Aspose.Cells | set custom icon for OLE object Aspose.Cells .NET | place OLE object at specific cell coordinates | save workbook with embedded audio Aspose
// Developer Intent: Insert a WAV audio file as an embedded OLE object at a defined cell location in an Excel workbook.
// Use Cases: Create interactive reports that play audio clips directly from spreadsheet cells. | Generate data sheets where each record includes an associated voice note embedded as an OLE object. | Distribute training manuals in Excel with built‑in sound cues, removing the need for external media files.
// AI Prompts: Show C# code that embeds a WAV file as an OLE object with a PNG icon using Aspose.Cells. | Provide an example that adds multiple audio OLE objects to different cells in the same workbook. | Explain how to extract and play the embedded WAV data from an OleObject after the workbook is saved.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to read a WAV file (and optional PNG icon), add an OleObject at given row/column coordinates, embed the audio data with SetEmbeddedObject, set a label, and save the workbook as XLSX using Aspose.Cells for .NET.
class InsertWavAsOleObject
{
    static void Main()
    {
        try
        {
            // Paths to the WAV file and an optional icon image to represent the OLE object.
            string wavFilePath = "sample.wav";          // Replace with your actual WAV file path.
            string iconImagePath = "audio_icon.png";    // Replace with an icon image path or leave empty.

            // Verify WAV file exists.
            if (!File.Exists(wavFilePath))
            {
                Console.WriteLine($"WAV file not found: {wavFilePath}");
                return;
            }

            // Read the WAV file bytes.
            byte[] wavData = File.ReadAllBytes(wavFilePath);

            // Read the icon image bytes; if the file does not exist, use an empty byte array.
            byte[] iconData = File.Exists(iconImagePath) ? File.ReadAllBytes(iconImagePath) : new byte[0];

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the position (row, column) and size (height, width in pixels) for the OLE object.
            int topRow = 5;          // Upper‑left row index (zero‑based).
            int leftColumn = 2;      // Upper‑left column index (zero‑based).
            int height = 100;        // Height in pixels.
            int width = 100;         // Width in pixels.

            // Add the OLE object placeholder using the icon image (or empty image).
            int oleIndex = worksheet.OleObjects.Add(topRow, leftColumn, height, width, iconData);

            // Retrieve the created OleObject.
            OleObject oleObject = worksheet.OleObjects[oleIndex];

            // Embed the WAV file data into the OleObject.
            oleObject.SetEmbeddedObject(
                linkToFile: false,
                objectData: wavData,
                sourceFileName: Path.GetFileName(wavFilePath),
                displayAsIcon: true,
                label: "Audio",
                updateIcon: false);

            // Optionally set the label that appears when the icon is hovered.
            oleObject.Label = "Sample Audio";

            // Save the workbook to an XLSX file.
            string outputPath = "WorkbookWithWavOleObject.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
