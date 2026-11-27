// Title: How to export only the active worksheet to an XML file with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export the currently active worksheet as XML | Save a single sheet to XML using Aspose.Cells .NET API | C# copy active worksheet to a new workbook and save as XML | Set ExportAllSheets false in Aspose.Cells and export active sheet | Generate XML from one worksheet with Aspose.Cells for .NET
// Tags: active sheet XML export using Aspose.Cells | single-sheet SaveFormat.Xml Aspose.Cells | worksheet copy to new workbook C# Aspose.Cells | ExportAllSheets property false Aspose.Cells | export specific worksheet as XML .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, populates the first worksheet, ensures it is the active sheet, copies that sheet into a new workbook after clearing the default sheet, and then saves the new workbook as an XML file using Aspose.Cells' SaveFormat.Xml.
class ExportActiveSheetToXml
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data to the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue(123);

            // Ensure the first sheet is the active sheet
            workbook.Worksheets.ActiveSheetIndex = 0;

            // Get the index of the active worksheet
            int activeIndex = workbook.Worksheets.ActiveSheetIndex;

            // Create a new workbook that will contain only the active worksheet
            Workbook activeOnlyWorkbook = new Workbook();
            // Remove the default sheet created with the new workbook
            activeOnlyWorkbook.Worksheets.Clear();

            // Add an empty worksheet to serve as the copy destination
            int destIndex = activeOnlyWorkbook.Worksheets.Add();
            Worksheet destSheet = activeOnlyWorkbook.Worksheets[destIndex];

            // Copy the active worksheet to the new workbook
            workbook.Worksheets[activeIndex].Copy(destSheet);

            // Define output file path
            string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "ActiveSheetExport.xml");

            // Export the active sheet to an XML file
            activeOnlyWorkbook.Save(outputFile, SaveFormat.Xml);
            Console.WriteLine($"Active sheet exported successfully to {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
