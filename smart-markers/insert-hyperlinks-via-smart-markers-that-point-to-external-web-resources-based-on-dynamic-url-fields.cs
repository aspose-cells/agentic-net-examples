using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartMarkerHyperlinkDemo
{
    // Simple data class containing a URL and optional display text
    public class LinkInfo
    {
        public string Url { get; set; }
        public string DisplayText { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Place a smart marker in cell A2 that will be replaced by the URL value
            sheet.Cells["A2"].PutValue("&=Url");

            // Prepare a data source with dynamic URLs
            List<LinkInfo> links = new List<LinkInfo>
            {
                new LinkInfo { Url = "https://www.aspose.com", DisplayText = "Aspose Home" },
                new LinkInfo { Url = "https://github.com/aspose-cells", DisplayText = "Aspose Cells GitHub" }
            };

            // Use WorkbookDesigner to process the smart marker
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Links", links);
            // The smart marker will be processed for the first item in the collection
            designer.Process();

            // After processing, the cell A2 now contains the URL string.
            // Add a hyperlink to that cell using the same URL.
            string url = sheet.Cells["A2"].StringValue;
            int hyperlinkIndex = sheet.Hyperlinks.Add("A2", 1, 1, url);
            Hyperlink hyperlink = sheet.Hyperlinks[hyperlinkIndex];
            // Set the display text (optional)
            hyperlink.TextToDisplay = links[0].DisplayText;
            // Set a screen tip (optional)
            hyperlink.ScreenTip = "Click to open " + links[0].DisplayText;

            // Save the workbook to an XLSX file
            workbook.Save("SmartMarkerHyperlinkDemo.xlsx");
        }
    }
}