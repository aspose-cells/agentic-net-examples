using System;
using Aspose.Cells;

namespace LoadVisibleSheetsDemo
{
    // Custom LoadFilter that loads only visible worksheets
    public class VisibleSheetsLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load the sheet only if it is visible
            if (sheet.IsVisible)
            {
                base.StartSheet(sheet);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string filePath = "SourceWorkbook.xlsx";

            // Create LoadOptions and assign the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new VisibleSheetsLoadFilter();

            // Load the workbook using the options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Display information about the loaded worksheets
            Console.WriteLine("Loaded Worksheets (visible only):");
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Console.WriteLine($"- {ws.Name} (Visible = {ws.IsVisible})");
            }

            // Optionally, save the workbook to verify the result
            workbook.Save("VisibleSheetsOnly.xlsx");
        }
    }
}