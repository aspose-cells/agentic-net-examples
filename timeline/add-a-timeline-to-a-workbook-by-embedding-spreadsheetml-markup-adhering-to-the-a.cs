using System;
using System.IO;
using Aspose.Cells;

class AddTimelineViaXml
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // SpreadsheetML markup that defines a timeline.
        // The markup follows the SpreadsheetML schema (http://schemas.openxmlformats.org/spreadsheetml/2006/main).
        // It creates a timeline named "SalesTimeline" that is linked to a pivot cache (id 1) and a range A1:A4.
        string timelineXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <timeline name=""SalesTimeline"">
    <timelineCache pivotCacheId=""1"">
      <timelineRange>Sheet1!A1:A4</timelineRange>
    </timelineCache>
  </timeline>
</worksheet>";

        // Load the XML into a memory stream so it can be passed to ImportXml.
        using (MemoryStream ms = new MemoryStream())
        using (StreamWriter writer = new StreamWriter(ms))
        {
            writer.Write(timelineXml);
            writer.Flush();
            ms.Position = 0;

            // Import the XML into the first worksheet (Sheet1) starting at cell A1 (row 0, column 0).
            wb.ImportXml(ms, "Sheet1", 0, 0);
        }

        // Save the workbook; the timeline is now embedded in the file.
        wb.Save("WorkbookWithTimeline.xlsx");
    }
}