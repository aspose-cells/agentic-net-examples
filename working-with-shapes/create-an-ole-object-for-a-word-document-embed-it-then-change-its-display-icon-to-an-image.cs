// Title: C# – Embed a Word Document as an OLE Object with a Custom Icon using Aspose.Cells
// Description: Demonstrates how to create a new Workbook, read a .docx file and a PNG icon, add an OLE object placeholder, embed the Word file, display it as an icon with a custom image, and save the result as an .xlsx workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | OLE object | embed Word document | custom icon | SetEmbeddedObject | Excel workbook | display as icon | Aspose.Cells for .NET | embedding files in Excel
// Common Searches: Aspose.Cells embed Word as OLE object C# | change OLE object icon in Excel using Aspose.Cells | SetEmbeddedObject custom icon example | how to add a Word file to Excel with a custom icon | Aspose.Cells OLE object display as icon
// Developer Intent: Add a Word file to an Excel sheet as an embedded OLE object and replace the default icon with a user‑provided image.
// Use Cases: Attach a company‑branded Word analysis to a financial report, showing a custom PNG icon for quick identification. | Create a template where users double‑click a custom‑icon OLE object to open an embedded policy document. | Distribute a spreadsheet package that bundles multiple Word manuals, each represented by a distinct custom icon.
// AI Prompts: Generate C# code with Aspose.Cells to embed a PDF as an OLE object and set a custom JPEG icon. | Explain the effect of linkToFile, displayAsIcon, and updateIcon parameters in SetEmbeddedObject. | Provide a step‑by‑step guide to replace the default OLE icon with any image after embedding a file using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsOleObjectDemo
{
    // Demonstrates how to create a new Workbook, read a .docx file and a PNG icon, add an OLE object placeholder, embed the Word file, display it as an icon with a custom image, and save the result as an .xlsx workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Paths to the Word document and the custom icon image
            string wordFilePath = "sample.docx";
            string iconFilePath = "icon.png";

            // Read the Word document bytes (the data to embed)
            byte[] wordData = File.ReadAllBytes(wordFilePath);

            // Read the icon image bytes (used as the display image)
            byte[] iconData = File.ReadAllBytes(iconFilePath);

            // Add an OLE object placeholder using the icon image.
            // Parameters: topRow, leftColumn, height (px), width (px), imageData
            int oleIndex = worksheet.OleObjects.Add(5, 2, 100, 100, iconData);

            // Retrieve the added OleObject
            OleObject oleObject = worksheet.OleObjects[oleIndex];

            // Embed the Word document data, display it as an icon, and set a label.
            // linkToFile = false (embed the data), displayAsIcon = true, updateIcon = false (keep custom icon)
            oleObject.SetEmbeddedObject(
                linkToFile: false,
                objectData: wordData,
                sourceFileName: Path.GetFileName(wordFilePath),
                displayAsIcon: true,
                label: "Word Document",
                updateIcon: false);

            // Ensure the object is shown as an icon (redundant but explicit)
            oleObject.DisplayAsIcon = true;

            // Save the workbook
            workbook.Save("WordOleObjectWithCustomIcon.xlsx");
        }
    }
}
