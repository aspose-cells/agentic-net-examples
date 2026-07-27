// Title: Hide an Embedded OLE Object in Aspose.Cells for .NET (C#) using OleObject.IsHidden
// Description: Demonstrates how to create a workbook, embed an OLE object (an Excel sheet) into a worksheet, set its ProgID, supply the object data from a memory stream, mark the object as hidden with OleObject.IsHidden, and save the file.
// Keywords: Aspose.Cells hide OLE object | OleObject.IsHidden C# | Aspose.Cells embedded OLE visibility | background OLE invisible Aspose | C# Aspose.Cells hide shape | Excel OLE object hidden programmatically | Aspose.Cells .NET example
// Common Searches: how to hide an OLE object in Aspose.Cells C# | set OleObject.IsHidden property Aspose.Cells | make embedded OLE objects invisible in generated Excel files | Aspose.Cells hide background objects programmatically | C# code to hide OLE objects in a worksheet
// Developer Intent: Hide an embedded OLE object so it does not appear in the worksheet.
// Use Cases: Add a supplemental Excel sheet as a hidden OLE object to keep the report layout clean | Store auxiliary data in hidden OLE objects that can be revealed by user interaction later | Create a template with placeholder OLE objects that are hidden until populated
// AI Prompts: Generate C# code that adds an OLE object to a worksheet using Aspose.Cells and sets OleObject.IsHidden to true | Explain how to loop through all OleObjects in a worksheet and hide those with a specific ProgID using Aspose.Cells | Provide a step‑by‑step guide to toggle the visibility of an OleObject after the workbook has been saved

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, embed an OLE object (an Excel sheet) into a worksheet, set its ProgID, supply the object data from a memory stream, mark the object as hidden with OleObject.IsHidden, and save the file.
class HideOleObjectsDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add an embedded OLE object without a custom icon (null image data)
            int oleIndex = sheet.OleObjects.Add(5, 2, 150, 200, null);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Set the ProgID (e.g., Excel sheet) and embed a simple workbook as data
            ole.ProgID = "Excel.Sheet";
            using (MemoryStream ms = new MemoryStream())
            {
                new Workbook().Save(ms, SaveFormat.Xlsx);
                ole.ObjectData = ms.ToArray();
            }

            // Hide the OLE object so it does not appear in the worksheet
            ole.IsHidden = true;

            // Save the workbook
            string outputPath = "HiddenOleObjectDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
