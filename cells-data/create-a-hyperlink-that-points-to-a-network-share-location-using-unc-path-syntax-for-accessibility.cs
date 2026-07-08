using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // UNC path to the network share (e.g., \\Server\Share\Folder\File.pdf)
        string uncPath = @"\\ServerName\SharedFolder\Document.pdf";

        // Add a hyperlink to cell B2 that points to the UNC location
        worksheet.Hyperlinks.Add("B2", 1, 1, uncPath);

        // Set the display text for the hyperlink (optional)
        Hyperlink hyperlink = worksheet.Hyperlinks[0];
        hyperlink.TextToDisplay = "Open Document";

        // Save the workbook
        workbook.Save("NetworkShareHyperlink.xlsx");
    }
}