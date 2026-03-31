using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace SlicerRefreshDemo
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook that contains slicers
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each slicer on the current worksheet
                foreach (Slicer slicer in sheet.Slicers)
                {
                    // Refresh the slicer so it reflects any changes in the underlying data
                    slicer.Refresh();
                }
            }

            // Save the workbook after refreshing slicers
            workbook.Save("output.xlsx");
        }
    }
}