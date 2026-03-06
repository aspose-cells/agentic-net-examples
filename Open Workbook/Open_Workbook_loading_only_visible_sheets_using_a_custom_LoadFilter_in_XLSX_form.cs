using System;
using Aspose.Cells;

namespace LoadVisibleSheetsDemo
{
    // Custom LoadFilter that loads only worksheets marked as visible
    public class VisibleSheetsLoadFilter : LoadFilter
    {
        // This method is called before each worksheet is loaded
        public override void StartSheet(Worksheet sheet)
        {
            // Load the sheet only if it is visible
            if (sheet.IsVisible)
            {
                // Use the default loading behavior for visible sheets
                base.StartSheet(sheet);
            }
            // If the sheet is hidden, do nothing – it will be skipped
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourceFile = "InputWorkbook.xlsx";

            // Configure load options with the custom filter
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new VisibleSheetsLoadFilter();

            // Load the workbook using the options – only visible sheets will be loaded
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Demonstrate that only visible sheets are present
            Console.WriteLine("Loaded worksheets:");
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Console.WriteLine($"- {ws.Name} (Visible = {ws.IsVisible})");
            }

            // Optionally save the filtered workbook to a new file
            string outputFile = "VisibleSheetsOnly.xlsx";
            workbook.Save(outputFile);

            // Clean up
            workbook.Dispose();
        }
    }
}