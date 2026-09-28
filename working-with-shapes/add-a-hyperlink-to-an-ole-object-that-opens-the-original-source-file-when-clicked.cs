// Title: Add a hyperlink to an embedded OLE object that opens the source file in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a Word document as an OLE object at cell B2 and set its Hyperlink.Address to the original file path with Aspose.Cells in C#. | Create an Excel workbook, embed any file as an OLE object, attach a clickable hyperlink that opens the source file, and save the workbook using Aspose.Cells for .NET. | Generate code that adds an OLE shape to a worksheet, assigns a file‑system hyperlink to the shape, and verifies the workbook saves correctly.
// Common Searches: Aspose.Cells C# add hyperlink to OLE shape in Excel | how to make OLE object open original document when clicked using Aspose.Cells | set Hyperlink.Address for OleObject in Aspose.Cells .NET example | embed Word file as OLE object with file link in Excel workbook using Aspose.Cells | C# Aspose.Cells save workbook with hyperlinked OLE object
// Tags: OleObject hyperlink Aspose.Cells | embed Word as OLE shape C# | Hyperlink.Address property usage | save workbook with OLE link | Excel OLE object file reference

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, embeds a Word document as an OLE object at cell B2, assigns the source file path as a hyperlink to the OLE shape, and saves the workbook as HyperlinkedOle.xlsx.
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

            // Path to the source file that will be embedded as an OLE object
            string sourceFilePath = @"C:\Temp\Sample.docx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourceFilePath))
            {
                Console.WriteLine($"Source file not found: {sourceFilePath}");
                return;
            }

            // Read the source file into a byte array (required by AddOleObject overload)
            byte[] oleData = File.ReadAllBytes(sourceFilePath);

            // Insert the OLE object at cell B2 (row index 1, column index 1)
            // Provide row/column offsets (0) and size in points
            OleObject oleObject = sheet.Shapes.AddOleObject(
                1,          // upperLeftRow
                1,          // upperLeftColumn
                0,          // upperLeftRowOffset
                0,          // upperLeftColumnOffset
                100,        // height
                200,        // width
                oleData);   // OLE data

            // Add a hyperlink to the OLE object that opens the original source file
            oleObject.Hyperlink.Address = sourceFilePath;

            // Ensure the output directory exists
            string outputPath = @"C:\Temp\HyperlinkedOle.xlsx";
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
