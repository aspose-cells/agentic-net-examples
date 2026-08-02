// Title: Batch add identical WAV OLE object to every worksheet with Aspose.Cells (C#)
// Description: Creates a new workbook, adds multiple sheets, verifies a WAV file and an icon, then embeds the same audio OLE object (displayed as a custom icon with AutoLoad) into each worksheet and saves the file as XLSX.
// Keywords: Aspose.Cells C# | embed WAV OLE object | Excel audio OLE | batch insert OLE | multiple worksheets | custom icon for OLE | AutoLoad OLE property | save workbook as XLSX | .NET Excel automation
// Common Searches: Aspose.Cells add same audio OLE to all sheets | C# embed WAV file as OLE in Excel workbook | batch insert OLE objects with custom icons using Aspose.Cells | set AutoLoad for OLE objects in .NET Excel | how to add audio OLE to every worksheet programmatically
// Developer Intent: Insert one WAV audio OLE object with a custom icon into every worksheet of an Excel workbook using Aspose.Cells.
// Use Cases: Generate a training workbook where each tab plays the same instructional audio. | Add a uniform audio cue to all sheets of a report for accessibility compliance. | Automate the preparation of template files that require a consistent embedded sound across pages.
// AI Prompts: Write C# code with Aspose.Cells that embeds a WAV file as an OLE object, shows a custom icon, enables AutoLoad, and repeats the insertion on all worksheets. | Refactor the example to extract the OLE insertion into a reusable method and add detailed error handling for missing audio or icon files. | Create step‑by‑step documentation for embedding audio OLE objects in Excel using Aspose.Cells, covering file requirements, icon usage, and property settings.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsBatchOle
{
    // Creates a new workbook, adds multiple sheets, verifies a WAV file and an icon, then embeds the same audio OLE object (displayed as a custom icon with AutoLoad) into each worksheet and saves the file as XLSX.
    public class AddWavOleToAllSheets
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add worksheets
                workbook.Worksheets.Add("Sheet1");
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Paths to resources
                string wavFilePath = "sample.wav";
                string iconFilePath = "audio_icon.png";

                // Verify files exist
                if (!File.Exists(wavFilePath))
                    throw new FileNotFoundException($"WAV file not found: {wavFilePath}");
                if (!File.Exists(iconFilePath))
                    throw new FileNotFoundException($"Icon file not found: {iconFilePath}");

                // Read file bytes
                byte[] wavData = File.ReadAllBytes(wavFilePath);
                byte[] iconData = File.ReadAllBytes(iconFilePath);

                // Add OLE object to each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int oleIndex = sheet.OleObjects.Add(5, 5, 100, 100, iconData);
                    OleObject ole = sheet.OleObjects[oleIndex];
                    ole.SetEmbeddedObject(
                        linkToFile: false,
                        objectData: wavData,
                        sourceFileName: wavFilePath,
                        displayAsIcon: true,
                        label: "Audio");
                    ole.AutoLoad = true;
                }

                // Save workbook
                workbook.Save("WorkbookWithWavOle.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            AddWavOleToAllSheets.Run();
        }
    }
}
