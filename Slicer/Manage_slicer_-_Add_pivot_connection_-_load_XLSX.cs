using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    class Program
    {
        static void Main()
        {
            // Paths to the source and destination files
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the workbook with options (enable parsing of pivot cached records if needed)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.ParsingPivotCachedRecords = true;
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Assume the first worksheet contains the data and a PivotTable
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the first PivotTable on the worksheet
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTable found on the worksheet.");
                return;
            }
            PivotTable pivot = sheet.PivotTables[0];

            // Determine a base field name from the PivotTable's BaseFields collection
            // (use the first base field as an example)
            string baseFieldName = pivot.BaseFields[0].Name;

            // Add a slicer linked to the PivotTable using the base field name
            // Row and column indices (0,0) place the slicer at the top‑left corner of the sheet
            int slicerIndex = sheet.Slicers.Add(pivot, 0, 0, baseFieldName);
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Explicitly add the PivotTable connection to the slicer (optional, but demonstrates the method)
            slicer.AddPivotConnection(pivot);

            // Refresh the slicer to ensure it reflects the current PivotTable data
            slicer.Refresh();

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}