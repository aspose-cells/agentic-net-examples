// Title: Embed a PDF as an OLE object displayed as an icon in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Insert a PDF file into the first worksheet of an existing Excel workbook as an OLE object that shows as an icon, using Aspose.Cells for C#. | Replace the default OLE icon with a custom .ico file for the embedded PDF and save the workbook to a new location with Aspose.Cells. | Read a PDF from disk, embed it as an OLE object at cell B2, set DisplayAsIcon = true, and export the modified workbook.
// Common Searches: Aspose.Cells C# embed PDF as OLE object icon | How to add an OLE object to an Excel worksheet with Aspose.Cells | Set custom icon for PDF OLE object in Aspose.Cells .NET | Save workbook after inserting OLE object using Aspose.Cells | Insert external file as OLE object in Excel via Aspose.Cells C#
// Tags: pdf oleobject embedding aspose.cells | oleobject icon display c# | worksheet oleobject insertion aspose.cells | custom oleobject icon aspose.cells | excel workbook save after oleobject aspose.cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an existing workbook, embeds a PDF file as an OLE object displayed as an icon on the first worksheet, optionally sets a custom icon, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string pdfPath = @"C:\Docs\sample.pdf";
            const string outputPath = "output.xlsx";

            // Verify required files exist
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input workbook not found: {inputPath}");
                return;
            }
            if (!File.Exists(pdfPath))
            {
                Console.WriteLine($"PDF file not found: {pdfPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Read PDF bytes for OLE embedding
            byte[] pdfData = File.ReadAllBytes(pdfPath);

            // Add an OLE object that points to the PDF file
            // OleObjects.Add returns the index of the newly added object
            int oleIndex = sheet.OleObjects.Add(1, 1, 100, 100, pdfData);

            // Retrieve the OleObject instance using the returned index
            OleObject ole = sheet.OleObjects[oleIndex];

            // Display the OLE object as an icon
            ole.DisplayAsIcon = true;

            // (Optional) Set a custom icon file or index
            // ole.IconFileName = @"C:\Path\To\Icon.ico";
            // ole.IconIndex = 0;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
