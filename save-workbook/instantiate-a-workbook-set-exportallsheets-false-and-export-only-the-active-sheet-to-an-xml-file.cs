// Title: How to export only the active worksheet to an XML file with Aspose.Cells for .NET
// AI Prompts: Generate C# code that saves only the active worksheet of a Workbook to an XML file using Aspose.Cells XmlSaveOptions and the SheetIndexes property. | Show how to set TxtSaveOptions.ExportAllSheets = false while exporting a specific worksheet to XML with Aspose.Cells.
// Common Searches: Aspose.Cells save only the active sheet as XML | Using XmlSaveOptions SheetIndexes to export a single worksheet in C# | Set ExportAllSheets false when saving a workbook with Aspose.Cells | C# export specific worksheet to XML using Aspose.Cells | Selective sheet export to XML in Aspose.Cells .NET
// Tags: XmlSaveOptions SheetIndexes selective export | export single worksheet to XML Aspose.Cells | TxtSaveOptions ExportAllSheets false example | active sheet XML save .NET | Aspose.Cells workbook selective sheet export

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program creates a workbook with two worksheets, marks the first sheet as active, configures XmlSaveOptions to include only that sheet via the SheetIndexes array, and saves it as ActiveSheetOnly.xml. It also demonstrates setting TxtSaveOptions.ExportAllSheets to false, which does not affect the XML export.
class Program
{
    static void Main()
    {
        // Instantiate a new workbook with two worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets.Add("Sheet2");

        // Populate some data in both sheets
        workbook.Worksheets[0].Cells["A1"].PutValue("Data in Sheet1");
        workbook.Worksheets[1].Cells["A1"].PutValue("Data in Sheet2");

        // Set the first worksheet as the active sheet
        workbook.Worksheets.ActiveSheetIndex = 0;

        // Set ExportAllSheets to false (property belongs to TxtSaveOptions,
        // included here as per the requirement even though it does not affect XML saving)
        TxtSaveOptions txtOptions = new TxtSaveOptions(SaveFormat.Csv);
        txtOptions.ExportAllSheets = false;

        // Configure XML save options to export only the active worksheet
        XmlSaveOptions xmlOptions = new XmlSaveOptions
        {
            // Export only the active sheet by specifying its index
            SheetIndexes = new int[] { workbook.Worksheets.ActiveSheetIndex }
        };

        // Export the active worksheet to an XML file
        workbook.Save("ActiveSheetOnly.xml", xmlOptions);
    }
}
