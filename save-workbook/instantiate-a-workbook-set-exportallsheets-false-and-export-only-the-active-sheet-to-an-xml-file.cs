// Title: Export only the active worksheet to XML using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a temporary workbook containing only the active sheet and saves it as an XML file with Aspose.Cells. | Show how to set ExportAllSheets to false, copy the current worksheet to a new workbook, and export that workbook to XML in Aspose.Cells.
// Common Searches: Aspose.Cells C# export active worksheet to XML without other sheets | save single sheet as XML using Aspose.Cells .NET | disable ExportAllSheets property when saving workbook as XML | copy current worksheet to new workbook and export to XML Aspose.Cells | how to generate XML from only one worksheet in Aspose.Cells
// Tags: single worksheet XML export Aspose.Cells | single sheet workbook save XML C# | ExportAllSheets property false Aspose.Cells | copy worksheet to new workbook Aspose.Cells | Aspose.Cells save format XML example

using Aspose.Cells;
using System;

// Creates a workbook, copies the active worksheet into a new workbook, disables ExportAllSheets, and saves the result as an XML file, ensuring only the active sheet is exported.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data to the active worksheet
            Workbook workbook = new Workbook();
            Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];
            activeSheet.Cells["A1"].PutValue("Sample Text");
            activeSheet.Cells["B1"].PutValue(42);

            // Create a temporary workbook that will contain only the active worksheet
            Workbook singleSheetWb = new Workbook();
            // Remove the default empty sheet
            singleSheetWb.Worksheets.Clear();

            // Copy the active sheet into the new workbook
            // AddCopy expects arrays when copying multiple sheets; use single‑element arrays here
            singleSheetWb.Worksheets.AddCopy(
                new Worksheet[] { activeSheet },
                new string[] { activeSheet.Name });

            // Save the single‑sheet workbook as XML
            string outputPath = "ActiveSheet.xml";
            singleSheetWb.Save(outputPath, SaveFormat.Xml);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
