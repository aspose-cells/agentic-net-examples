// Title: How to embed a Word .docx file as an OLE object and show it as an icon in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads a .docx file, adds it as an OLE object to a specific cell range in an Aspose.Cells worksheet, and configures the object to display as an icon. | Show how to embed a Word document into an Excel workbook with Aspose.Cells, enable the icon view, and export the file as .xlsx. | Describe the current Aspose.Cells limitation on assigning a custom icon to an OLE object and propose an alternative method to show a custom image.
// Common Searches: asp.net embed word .docx as ole object in excel using aspose.cells | c# set oleobject displayasicon true with aspose.cells | aspose.cells custom icon for ole object not supported | programmatically add ole object to worksheet c# | save excel workbook with embedded word file aspose.cells example
// Tags: add ole object to worksheet c# | embed word document as ole object aspose.cells | display ole object as icon aspose.cells | custom icon limitation oleobject aspose.cells | save workbook with embedded oleobject c#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for OleObject

// Creates a new workbook, reads a .docx file into a byte array, adds it as an OLE object at row 5 column 2, sets DisplayAsIcon = true (default Excel icon), and saves the workbook as .xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define file paths
            string wordFilePath = @"C:\Docs\SampleDocument.docx";
            string iconFilePath = @"C:\Icons\CustomIcon.ico";
            string outputPath = @"C:\Output\WorkbookWithOleObject.xlsx";

            // Verify required files exist
            if (!File.Exists(wordFilePath))
                throw new FileNotFoundException("Word document not found.", wordFilePath);
            if (!File.Exists(iconFilePath))
                throw new FileNotFoundException("Icon file not found.", iconFilePath);

            // Read the Word document into a byte array (required by OleObjects.Add)
            byte[] oleData = File.ReadAllBytes(wordFilePath);

            // Add the OLE object to the worksheet (row 5, column 2, width 200, height 100)
            int oleIndex = sheet.OleObjects.Add(5, 2, 200, 100, oleData);

            // Retrieve the OLE object instance
            OleObject oleObject = sheet.OleObjects[oleIndex];

            // Display the OLE object as an icon
            oleObject.DisplayAsIcon = true;

            // NOTE: In recent Aspose.Cells versions the properties for setting a custom icon
            // (IconFile / IconIndex) are not available. The default Excel icon will be used.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook with the embedded OLE object
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully at: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
