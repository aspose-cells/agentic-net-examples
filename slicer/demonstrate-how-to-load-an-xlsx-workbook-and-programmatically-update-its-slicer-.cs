using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        string inputPath = "InputWithSlicer.xlsx";

        // Ensure the input file exists; create an empty workbook if it does not.
        if (!File.Exists(inputPath))
        {
            Workbook emptyWb = new Workbook();
            emptyWb.Worksheets[0].Name = "Sheet1";
            emptyWb.Save(inputPath);
        }

        // Load the workbook.
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];

        // Access the slicer collection on this worksheet.
        SlicerCollection slicers = sheet.Slicers;

        // If there is at least one slicer, update its settings.
        if (slicers.Count > 0)
        {
            Slicer slicer = slicers[0];
            slicer.Caption = "Updated Slicer Caption";
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            slicer.NumberOfColumns = 2;
            slicer.LockedPosition = false;
            slicer.ShowAllItems = true;
            slicer.Refresh();
        }

        // Save the workbook with the modified slicer settings.
        workbook.Save("OutputWithUpdatedSlicer.xlsx");
    }
}