// Title: C# – Embed a Word Document as an OLE Object with a Custom Icon using Aspose.Cells for .NET
// Description: Creates a new workbook, adds an OLE placeholder with a PNG image, embeds a .docx file, forces the object to display as an icon, assigns the custom image as the icon, sets the Word ProgID, and saves the result as an .xlsx file.
// Keywords: Aspose.Cells OLE object | embed Word document Excel | custom OLE icon C# | SetEmbeddedObject Aspose.Cells | display as icon Excel OLE | ProgID Word OleObject | C# Excel embed docx
// Common Searches: Aspose.Cells embed docx as OLE object | change OLE object icon in Excel with Aspose.Cells | C# add OLE object with custom image placeholder | set custom display icon for embedded Word file | update ProgID for Word OLE object Aspose.Cells
// Developer Intent: Embed a Word file into an Excel worksheet as an OLE object and replace the default icon with a user‑provided PNG image.
// Use Cases: Financial reports that link to contract documents via recognizable icons. | Template workbooks where users click a custom icon to open policy PDFs stored inside the sheet. | Automated packaging of multiple reference files, each represented by a distinct PNG icon in a single spreadsheet.
// AI Prompts: Generate C# code with Aspose.Cells to embed a PDF as an OLE object and assign a custom PNG icon. | Show how to update the icon of an existing OLE object in a saved workbook using Aspose.Cells. | Provide best‑practice error handling for embedding large Word files as OLE objects in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsOleObjectExample
{
    // Creates a new workbook, adds an OLE placeholder with a PNG image, embeds a .docx file, forces the object to display as an icon, assigns the custom image as the icon, sets the Word ProgID, and saves the result as an .xlsx file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths to the source Word document and the custom icon image
                string wordFilePath = "sample.docx";          // Word document to embed
                string iconImagePath = "word_icon.png";       // Image to use as display icon

                // Verify that required files exist
                if (!File.Exists(wordFilePath))
                {
                    Console.WriteLine($"Word file not found: {wordFilePath}");
                    return;
                }

                if (!File.Exists(iconImagePath))
                {
                    Console.WriteLine($"Icon image not found: {iconImagePath}");
                    return;
                }

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Read the icon image bytes – this will be used as the placeholder picture for the OLE object
                byte[] iconImageBytes = File.ReadAllBytes(iconImagePath);

                // Add an OLE object placeholder to the worksheet (initial image is the icon)
                // Parameters: topRow, leftColumn, height (px), width (px), imageData
                int oleIndex = worksheet.OleObjects.Add(5, 2, 100, 100, iconImageBytes);

                // Retrieve the added OleObject
                OleObject oleObject = worksheet.OleObjects[oleIndex];

                // Read the Word document bytes to embed
                byte[] wordData = File.ReadAllBytes(wordFilePath);

                // Embed the Word document and set it to display as an icon
                // linkToFile = false (embed), displayAsIcon = true, updateIcon = false (keep custom icon)
                oleObject.SetEmbeddedObject(
                    linkToFile: false,
                    objectData: wordData,
                    sourceFileName: Path.GetFileName(wordFilePath),
                    displayAsIcon: true,
                    label: "Word Document",
                    updateIcon: false);

                // Ensure the object is shown as an icon
                oleObject.DisplayAsIcon = true;

                // Set the custom icon image source (the same image used when adding the placeholder)
                oleObject.ImageSourceFullName = iconImagePath;

                // Optionally set the ProgID for Word documents
                oleObject.ProgID = "Word.Document.12";

                // Save the workbook
                string outputPath = "WordOleObjectWithCustomIcon.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
