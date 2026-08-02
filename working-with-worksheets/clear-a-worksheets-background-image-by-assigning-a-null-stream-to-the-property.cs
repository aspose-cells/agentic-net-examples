// Title: Clear a Worksheet Background Image with Aspose.Cells for .NET (C#)
// Description: This C# example creates a new Workbook, accesses the first Worksheet, removes its background image by assigning null to the BackgroundImage property, and saves the file as an XLSX document.
// Keywords: Aspose.Cells | C# | BackgroundImage | clear worksheet background | remove worksheet picture | set BackgroundImage null | Aspose.Cells .NET | .xlsx save | Excel background image removal | GitHub Aspose.Cells example
// Common Searches: How to delete a worksheet background image using Aspose.Cells for .NET | Aspose.Cells clear background picture C# code | Set BackgroundImage to null in Aspose.Cells | Remove Excel worksheet background with Aspose.Cells | Clear worksheet background without reloading workbook
// Developer Intent: Remove the existing background image from a worksheet.
// Use Cases: Strip default background images from template workbooks before data insertion | Produce printer‑friendly reports by eliminating background graphics | Reset worksheet styling when reusing the same sheet for different datasets
// AI Prompts: Provide a C# snippet that clears a worksheet's background image using Aspose.Cells and saves the workbook. | Show how to set the BackgroundImage property to null for a specific worksheet index in an Aspose.Cells workbook. | Demonstrate how to confirm that a worksheet no longer contains a background image after clearing it with Aspose.Cells.

using System;
using Aspose.Cells;

// This C# example creates a new Workbook, accesses the first Worksheet, removes its background image by assigning null to the BackgroundImage property, and saves the file as an XLSX document.
class ClearWorksheetBackgroundImage
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Clear the background image by setting the property to null
        worksheet.BackgroundImage = null;

        // Save the workbook
        workbook.Save("ClearedBackground.xlsx", SaveFormat.Xlsx);
    }
}
