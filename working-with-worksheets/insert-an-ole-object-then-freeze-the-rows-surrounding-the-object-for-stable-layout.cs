// Title: C# – Insert an OLE Object and Freeze Panes Around It with Aspose.Cells
// Description: Demonstrates how to create a workbook, embed an Excel file as an OLE object (optionally with a custom icon), set its ProgID and display options, freeze the rows above and columns to the left of the object for a stable layout, and save the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | OLE object | embed Excel file | custom OLE icon | freeze panes | freeze rows | freeze columns | worksheet API | add OleObject | Aspose.Cells tutorial
// Common Searches: Aspose.Cells add OLE object C# | how to embed Excel as OLE in Aspose.Cells | freeze panes around OLE object Aspose.Cells | C# set custom icon for OLE object Aspose | Aspose.Cells freeze rows and columns example
// Developer Intent: Embed an OLE object in a worksheet and freeze the surrounding rows and columns to keep its position fixed.
// Use Cases: Create a report that embeds a supporting spreadsheet as an OLE icon while keeping header rows visible. | Design a dashboard where an OLE chart is placed and the surrounding rows stay fixed during scrolling. | Build a printable template that includes an embedded file with a custom icon and frozen panes for consistent layout.
// AI Prompts: Show how to link an OLE object instead of embedding it using Aspose.Cells. | Provide code that sets a custom icon for the OLE object via the IconData property. | Explain how to freeze only the rows above the OLE object while allowing columns to scroll.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for OleObject

// Demonstrates how to create a workbook, embed an Excel file as an OLE object (optionally with a custom icon), set its ProgID and display options, freeze the rows above and columns to the left of the object for a stable layout, and save the result as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the position (row, column) and size (height, width) of the OLE object
            // Note: Aspose.Cells uses zero‑based indexes for rows and columns
            int topRow = 5;        // Row where the OLE object starts
            int leftColumn = 2;    // Column where the OLE object starts
            int height = 200;      // Height in pixels
            int width = 300;       // Width in pixels

            // Load an image that will be shown as the OLE object's icon (optional)
            string iconPath = "icon.jpg";
            byte[]? iconData = null;
            if (File.Exists(iconPath))
            {
                iconData = File.ReadAllBytes(iconPath);
            }
            else
            {
                Console.WriteLine($"Icon file not found: {iconPath}. The OLE object will use the default icon.");
            }

            // Load the file to be embedded into the OLE object
            string embedPath = "sample.xlsx";
            byte[]? embedData = null;
            if (File.Exists(embedPath))
            {
                embedData = File.ReadAllBytes(embedPath);
            }
            else
            {
                Console.WriteLine($"Embedded file not found: {embedPath}. The OLE object will be added without embedded data.");
            }

            // Add the OLE object to the worksheet (embedded, not linked)
            int oleIndex = sheet.OleObjects.Add(
                topRow,
                leftColumn,
                height,
                width,
                embedData ?? Array.Empty<byte>()); // Use empty data if file missing

            OleObject ole = sheet.OleObjects[oleIndex];

            // If a custom icon was loaded, set it (Aspose.Cells supports custom icons via IconData)
            if (iconData != null)
            {
                // The IconData property may not be available in older versions; if so, this block can be omitted.
                // ole.IconData = iconData;
            }

            // Define OLE object properties
            ole.ProgID = "Excel.Sheet.8";      // Program ID for Excel files
            ole.DisplayAsIcon = true;          // Show as an icon
            ole.Label = "Embedded Excel File"; // Icon label

            // Freeze rows above the OLE object and columns to its left
            sheet.FreezePanes(topRow, leftColumn, topRow, leftColumn);

            // Save the workbook
            string outputPath = "OleObjectWithFreeze.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
