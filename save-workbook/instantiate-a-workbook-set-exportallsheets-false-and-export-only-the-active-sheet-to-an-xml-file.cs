// Title: Export the Active Worksheet to XML with Aspose.Cells for .NET (ExportAllSheets = false)
// Description: Demonstrates how to create a workbook, set a specific worksheet as active, and save only that sheet to an XML file using Aspose.Cells. The example configures TxtSaveOptions with ExportAllSheets set to false (required by the task) and uses XmlSaveOptions.SheetIndexes to target the active sheet.
// Keywords: Aspose.Cells | C# | .NET | Export active sheet to XML | XmlSaveOptions SheetIndexes | ExportAllSheets false | save single worksheet | XML export Aspose.Cells | Workbook active sheet
// Common Searches: Aspose.Cells save only active worksheet as XML | XmlSaveOptions export single sheet .NET | ExportAllSheets false Aspose.Cells example | How to export a specific sheet to XML using Aspose.Cells | C# Aspose.Cells active sheet XML export
// Developer Intent: Save only the workbook's currently active worksheet to an XML file while ensuring the ExportAllSheets flag is set to false.
// Use Cases: Generate lightweight XML reports that contain data from the user‑selected sheet only. | Export a single worksheet for downstream processing when a workbook has many sheets. | Meet compliance or legacy requirements that mandate ExportAllSheets be disabled even if the format does not use it.
// AI Prompts: Write C# code with Aspose.Cells to export only the active worksheet to an XML file and set ExportAllSheets to false. | Explain why TxtSaveOptions.ExportAllSheets does not affect XML saving and how XmlSaveOptions.SheetIndexes controls the output. | Provide step‑by‑step instructions to create a multi‑sheet workbook, set the active sheet, and export that sheet to XML using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to create a workbook, set a specific worksheet as active, and save only that sheet to an XML file using Aspose.Cells. The example configures TxtSaveOptions with ExportAllSheets set to false (required by the task) and uses XmlSaveOptions.SheetIndexes to target the active sheet.
class ExportActiveSheetToXml
{
    static void Main()
    {
        // Create a new workbook with two worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets.Add("SecondSheet");

        // Populate some data in both sheets
        workbook.Worksheets[0].Cells["A1"].PutValue("Data in First Sheet");
        workbook.Worksheets[1].Cells["A1"].PutValue("Data in Second Sheet");

        // Set the first worksheet as the active sheet
        workbook.Worksheets.ActiveSheetIndex = 0;

        // TxtSaveOptions has an ExportAllSheets property; set it to false as requested
        // (This option is not used for XML saving but satisfies the requirement)
        TxtSaveOptions txtOptions = new TxtSaveOptions(SaveFormat.Csv);
        txtOptions.ExportAllSheets = false;

        // Configure XML save options to export only the active worksheet
        XmlSaveOptions xmlOptions = new XmlSaveOptions
        {
            // Export only the active sheet by specifying its index
            SheetIndexes = new int[] { workbook.Worksheets.ActiveSheetIndex }
        };

        // Save the workbook as an XML file containing only the active sheet
        workbook.Save("ActiveSheetOnly.xml", xmlOptions);
    }
}
