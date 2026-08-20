// Title: How to add a hyperlink to a linked OLE object in an Excel worksheet using Aspose.Cells for .NET (C#)
// Description: Demonstrates creating a workbook, inserting a linked OLE object that references an existing file, assigning a Hyperlink (address and ScreenTip) to the OleObject, and saving the workbook so the object opens the original document on click.
// Keywords: Aspose.Cells | C# | .NET | OLE object hyperlink | linked OLE | Excel hyperlink | open source file | Hyperlink.Address | ScreenTip | worksheet OleObjects | Aspose.Cells API
// Common Searches: Aspose.Cells add hyperlink to OLE object C# | set hyperlink on linked OLE object Aspose.Cells | open original file from OLE object Excel .NET | OleObject Hyperlink example Aspose | C# code to link OLE object with hyperlink
// Developer Intent: Create a clickable OLE object that launches its source file when the user clicks it.
// Use Cases: Embed a Word contract as a linked OLE object in a financial report and provide a direct hyperlink for reviewers. | Generate a project tracker where each task includes a linked PDF OLE object with a hyperlink for quick document access. | Automate audit workbooks that insert source spreadsheets as linked OLE objects, each equipped with a hyperlink for auditors to open the original files.
// AI Prompts: Write C# code with Aspose.Cells to insert a linked OLE object and set a Hyperlink that opens the source file when clicked. | Explain how to test that an OleObject hyperlink works after saving the workbook with Aspose.Cells. | Provide a sample that adds a ScreenTip to an OLE object's hyperlink and saves the workbook to a custom path.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectHyperlinkDemo
{
    // Demonstrates creating a workbook, inserting a linked OLE object that references an existing file, assigning a Hyperlink (address and ScreenTip) to the OleObject, and saving the workbook so the object opens the original document on click.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the source file that will be linked and opened on click
                string sourceFilePath = @"C:\Temp\SampleDocument.docx";

                // Ensure the source file exists before proceeding
                if (!File.Exists(sourceFilePath))
                {
                    Console.WriteLine($"Source file not found: {sourceFilePath}");
                    return;
                }

                // Read the source file as a byte array (the OLE object data)
                byte[] oleData = File.ReadAllBytes(sourceFilePath);

                // Add a linked OLE object to the worksheet
                int oleIndex = worksheet.OleObjects.Add(
                    topRow: 5,
                    leftColumn: 2,
                    height: 200,
                    width: 300,
                    imageData: oleData,
                    linkedFile: sourceFilePath);

                // Retrieve the added OLE object
                OleObject oleObject = worksheet.OleObjects[oleIndex];

                // Set hyperlink for the OLE object (read‑only property, modify its members)
                Hyperlink hyperlink = oleObject.Hyperlink;
                hyperlink.Address = sourceFilePath;
                hyperlink.ScreenTip = "Open linked document";

                // Save the workbook
                string outputPath = "OleObjectWithHyperlink.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
