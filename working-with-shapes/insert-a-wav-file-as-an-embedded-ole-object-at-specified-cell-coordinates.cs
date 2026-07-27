// Title: Embed a WAV audio file as an OLE object in a specific Excel cell using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, positions an OLE placeholder at given row/column coordinates, optionally loads a PNG icon, reads a WAV file, embeds the audio with SetEmbeddedObject (linkToFile = false, displayAsIcon = false), and saves the file as an .xlsx document.
// Keywords: Aspose.Cells | C# | embed WAV | OLE object | Excel audio | custom OLE icon | SetEmbeddedObject | worksheet OLE | programmatic Excel audio
// Common Searches: Aspose.Cells embed wav in Excel C# | Add audio OLE object to a cell with Aspose.Cells | SetEmbeddedObject displayAsIcon false Aspose | Load custom icon for OLE object Aspose.Cells | Insert WAV file as OLE in .NET Excel
// Developer Intent: Insert a WAV file as an embedded OLE object at a defined cell range in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Generate a sales report where each product row includes an audio description embedded directly in the sheet. | Build training worksheets that play a voice‑over clip when the user opens the OLE object in a specific cell. | Automate a catalog that attaches a short sound preview to each item by placing WAV files at predetermined cell locations.
// AI Prompts: Show how to embed an MP3 file as an OLE object with Aspose.Cells for .NET. | Add comprehensive error handling for missing icon or audio files when creating OLE objects in C#. | Demonstrate how to extract and play an embedded WAV file from an Excel workbook using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsOleObjectWavDemo
{
    // Creates a new workbook, positions an OLE placeholder at given row/column coordinates, optionally loads a PNG icon, reads a WAV file, embeds the audio with SetEmbeddedObject (linkToFile = false, displayAsIcon = false), and saves the file as an .xlsx document.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the cell coordinates where the OLE object will be placed
                int topRow = 5;      // zero‑based row index
                int leftColumn = 2;  // zero‑based column index
                int height = 200;    // height in pixels
                int width = 300;     // width in pixels

                // Load the icon image that will be displayed for the OLE object
                byte[] iconData;
                const string iconPath = "icon.png";
                if (File.Exists(iconPath))
                {
                    iconData = File.ReadAllBytes(iconPath);
                }
                else
                {
                    // Use an empty byte array if the icon file is missing
                    iconData = new byte[0];
                    Console.WriteLine($"Warning: Icon file '{iconPath}' not found. Using empty icon.");
                }

                // Add an empty OLE object to the worksheet (imageData is the icon)
                int oleIndex = worksheet.OleObjects.Add(topRow, leftColumn, height, width, iconData);
                OleObject oleObject = worksheet.OleObjects[oleIndex];

                // Load the WAV file that will be embedded
                byte[] wavData;
                const string wavPath = "sample.wav";
                if (File.Exists(wavPath))
                {
                    wavData = File.ReadAllBytes(wavPath);
                }
                else
                {
                    // If the WAV file is missing, skip embedding but still create the OLE placeholder
                    wavData = new byte[0];
                    Console.WriteLine($"Warning: WAV file '{wavPath}' not found. OLE object will be empty.");
                }

                // Embed the WAV file into the OLE object
                // linkToFile = false (embed the data), displayAsIcon = false (show actual content when opened)
                oleObject.SetEmbeddedObject(
                    linkToFile: false,
                    objectData: wavData,
                    sourceFileName: wavPath,
                    displayAsIcon: false,
                    label: "Audio Clip");

                // Optional: set a label and ensure the object is not displayed as an icon
                oleObject.Label = "Audio Clip";
                oleObject.DisplayAsIcon = false;

                // Save the workbook
                const string outputPath = "WorkbookWithWavOleObject.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
