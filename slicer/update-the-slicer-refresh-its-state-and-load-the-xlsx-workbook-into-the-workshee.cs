using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerRefreshDemo
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing XLSX workbook from file
            string inputFile = "InputWorkbook.xlsx";
            Workbook workbook = new Workbook(inputFile); // uses Workbook(string) constructor

            // Assume the slicer is placed on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Access the slicer collection of the worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Ensure there is at least one slicer present
            if (slicers.Count > 0)
            {
                // Get the first slicer
                Slicer slicer = slicers[0];

                // Update slicer properties as needed (example: change the caption)
                slicer.Caption = "Updated Slicer Caption";

                // Refresh the slicer – this also refreshes the underlying PivotTable(s)
                slicer.Refresh();
            }
            else
            {
                Console.WriteLine("No slicers found in the worksheet.");
            }

            // Save the modified workbook to a new file
            string outputFile = "OutputWorkbook.xlsx";
            workbook.Save(outputFile); // uses Workbook.Save(string) method
        }
    }
}