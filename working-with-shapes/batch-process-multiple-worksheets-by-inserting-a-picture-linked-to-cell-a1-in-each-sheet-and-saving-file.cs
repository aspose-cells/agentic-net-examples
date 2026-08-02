// Title: C# – Batch insert a linked picture into cell A1 of every worksheet with Aspose.Cells
// Description: Demonstrates how to create or load a workbook, clear default sheets, add multiple worksheets, and loop through each sheet to place a linked picture (from a URL or local path) at cell A1 using Shapes.AddLinkedPicture. The workbook is then saved as an XLSX file. Ideal for .NET developers needing consistent branding or external images across all sheets.
// Keywords: Aspose.Cells C# linked picture | AddLinkedPicture multiple worksheets | batch insert image Excel .NET | insert picture at A1 Aspose.Cells | Excel workbook logo automation | C# Aspose.Cells example GitHub | programmatic Excel image insertion | linked picture from URL Excel
// Common Searches: How to add the same linked image to all worksheets using Aspose.Cells | Batch insert picture at A1 in each sheet C# | Aspose.Cells add linked picture from URL to multiple sheets | Insert external image into every Excel worksheet programmatically | Aspose.Cells Shapes.AddLinkedPicture loop example
// Developer Intent: Insert an identical linked picture into cell A1 of every worksheet and save the workbook.
// Use Cases: Apply a web‑hosted logo to all sheets of a report so the image updates automatically when the source changes. | Create multi‑sheet templates that share a common diagram or chart without duplicating the file. | Automate branding by adding a linked image to each worksheet of newly generated workbooks.
// AI Prompts: Generate C# code that loads an existing workbook and adds a linked picture from a given URL to cell A1 of each worksheet, preserving existing content. | Show an Aspose.Cells example that iterates over all worksheets, inserts a linked picture with custom width and height, and includes robust exception handling. | Explain how to replace the linked picture on selected sheets while leaving other sheets unchanged using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.IO;

// Demonstrates how to create or load a workbook, clear default sheets, add multiple worksheets, and loop through each sheet to place a linked picture (from a URL or local path) at cell A1 using Shapes.AddLinkedPicture. The workbook is then saved as an XLSX file. Ideal for .NET developers needing consistent branding or external images across all sheets.
class BatchInsertLinkedPicture
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing file)
            Workbook workbook = new Workbook();

            // Remove the default worksheet to avoid duplicate names
            workbook.Worksheets.Clear();

            // Add sample worksheets with unique names
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // URL or local path of the image to be linked
            string imagePath = "https://example.com/sample.jpg";

            // Insert a linked picture at cell A1 (row 0, column 0) in every worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Shapes.AddLinkedPicture(0, 0, 100, 100, imagePath);
            }

            // Save the workbook with the inserted linked pictures
            string outputFile = "BatchLinkedPictures.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully: {Path.GetFullPath(outputFile)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
