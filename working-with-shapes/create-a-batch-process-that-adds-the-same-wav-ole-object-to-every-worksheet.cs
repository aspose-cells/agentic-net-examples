// Title: Batch Add a WAV OLE Object to All Worksheets with Aspose.Cells for .NET
// Description: Loads an existing workbook, reads a WAV file once, and inserts it as an OLE object on every sheet at cell E5 (100 × 100 px). Sets ProgID, displays the audio as an icon, and saves the updated file.
// Keywords: Aspose.Cells OLE object | embed audio Excel C# | batch add OLE to worksheets | WAV file as icon Aspose.Cells | C# Excel automation audio
// Common Searches: Aspose.Cells add same audio OLE to each worksheet | C# batch embed WAV file in Excel workbook | how to display audio icon in Excel using Aspose | programmatically insert OLE objects on all sheets | set ProgID for WAV OLE object Aspose.Cells
// Developer Intent: Programmatically embed a single WAV file as an OLE object on every worksheet of an existing Excel workbook.
// Use Cases: Create a training workbook where each sheet plays the same instruction audio. | Generate multi‑sheet reports that include a uniform audio cue for navigation. | Automate presentation decks stored in Excel, adding a consistent audio clip to each slide.
// AI Prompts: Show how to change the icon image for the embedded WAV OLE object in Aspose.Cells. | Provide code to place the OLE object at a variable cell address instead of a fixed location. | Explain techniques for minimizing memory usage when embedding large audio files across many worksheets.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace BatchOleObjectAdder
{
    // Loads an existing workbook, reads a WAV file once, and inserts it as an OLE object on every sheet at cell E5 (100 × 100 px). Sets ProgID, displays the audio as an icon, and saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string wavPath = "sample.wav";
                const string outputPath = "output.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input workbook not found: {inputPath}");
                    return;
                }

                // Verify that the WAV file exists
                if (!File.Exists(wavPath))
                {
                    Console.WriteLine($"WAV file not found: {wavPath}");
                    return;
                }

                // Load the workbook to which the OLE objects will be added
                Workbook workbook = new Workbook(inputPath);

                // Read the WAV file once – the same data will be used for every worksheet
                byte[] wavData = File.ReadAllBytes(wavPath);

                // Iterate through all worksheets and add the WAV OLE object
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Add an OLE object placeholder (image data) at row 5, column 5 with size 100x100 pixels
                    int oleIndex = sheet.OleObjects.Add(5, 5, 100, 100, wavData);

                    // Retrieve the newly added OLE object to set additional properties
                    OleObject ole = sheet.OleObjects[oleIndex];

                    // Set the actual embedded WAV data
                    ole.ObjectData = wavData;

                    // Optional: define the ProgID for a WAV file and display it as an icon
                    ole.ProgID = "WavAudio";
                    ole.DisplayAsIcon = true;
                    ole.Label = "Audio Clip";
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
