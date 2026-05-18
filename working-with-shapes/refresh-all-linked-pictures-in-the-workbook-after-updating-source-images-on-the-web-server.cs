using System;
using Aspose.Cells;

class RefreshLinkedPictures
{
    static void Main()
    {
        // Load the workbook that contains linked pictures.
        // The linked pictures point to images hosted on a web server.
        Workbook workbook = new Workbook("input.xlsx");

        // Refresh all external links in the workbook.
        // This includes linked pictures; Aspose.Cells updates the picture data
        // from the source URLs when RefreshAll is invoked.
        workbook.Worksheets.RefreshAll();

        // Save the workbook after the refresh operation.
        workbook.Save("output.xlsx");
    }
}