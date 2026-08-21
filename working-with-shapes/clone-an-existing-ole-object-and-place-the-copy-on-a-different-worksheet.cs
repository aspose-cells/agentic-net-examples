// Title: Clone an OLE Object to a Different Worksheet with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to duplicate an existing OLE object in a workbook by reading its binary data, creating a new worksheet, and adding a cloned OleObject with identical position, size, and visual settings using Aspose.Cells for .NET. The workbook is saved as ClonedOleObject.xlsx.
// Keywords: Aspose.Cells | C# | OLE object clone | copy OLE object worksheet | OleObjects.Add | Aspose.Cells.Drawing | embedded Excel file | CloneOleObjectDemo | Excel automation | save workbook Xlsx
// Common Searches: clone OLE object Aspose.Cells C# | copy embedded Excel file to another sheet .NET | duplicate OleObject across worksheets | Aspose.Cells how to copy OLE object | C# code to clone OLE object in Excel
// Developer Intent: Create an identical copy of a source OleObject and insert it into a new worksheet within the same workbook.
// Use Cases: Insert the same embedded document on multiple report sections without re‑embedding the file. | Provide language‑specific sheets that share a common OLE chart or spreadsheet. | Maintain consistent icon display and label when reusing OLE objects across worksheets.
// AI Prompts: Write C# code that clones an OleObject from one worksheet to another using Aspose.Cells, preserving size and display properties. | Explain step‑by‑step how to copy an OLE object's binary data and visual attributes to a new sheet with Aspose.Cells for .NET. | Suggest robust error handling for missing source file or invalid OLE data when cloning objects in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to duplicate an existing OLE object in a workbook by reading its binary data, creating a new worksheet, and adding a cloned OleObject with identical position, size, and visual settings using Aspose.Cells for .NET. The workbook is saved as ClonedOleObject.xlsx.
class CloneOleObjectDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -----------------------------------------------------------------
        // Source worksheet: add an OLE object that we will clone later
        // -----------------------------------------------------------------
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "Source";

        // Load binary data for the OLE object (e.g., an existing Excel file)
        // Adjust the path to a valid file on your system
        byte[] oleData = File.ReadAllBytes("sample.xlsx");

        // Add the OLE object to the source sheet
        // Parameters: topRow, leftColumn, height (px), width (px), object data
        int oleIndex = sourceSheet.OleObjects.Add(5, 2, 200, 300, oleData);
        OleObject sourceOle = sourceSheet.OleObjects[oleIndex];

        // Set some optional properties on the source OLE object
        sourceOle.DisplayAsIcon = true;
        sourceOle.Label = "Sample OLE";

        // -----------------------------------------------------------------
        // Target worksheet: clone the OLE object here
        // -----------------------------------------------------------------
        Worksheet targetSheet = workbook.Worksheets.Add("Clone");

        // Clone the OLE object by adding a new one with the same binary data
        // and copying position/size from the source object
        int clonedIndex = targetSheet.OleObjects.Add(
            sourceOle.UpperLeftRow,          // same top row
            sourceOle.UpperLeftColumn,       // same left column
            sourceOle.Height,                // same height
            sourceOle.Width,                 // same width
            sourceOle.ObjectData);           // same embedded data

        OleObject clonedOle = targetSheet.OleObjects[clonedIndex];

        // Copy additional visual properties to keep the clone identical
        clonedOle.DisplayAsIcon = sourceOle.DisplayAsIcon;
        clonedOle.Label = sourceOle.Label;

        // -----------------------------------------------------------------
        // Save the workbook with the cloned OLE object
        // -----------------------------------------------------------------
        workbook.Save("ClonedOleObject.xlsx", SaveFormat.Xlsx);
    }
}
