// Title: Exclude Hidden Worksheets on Load with Aspose.Cells .NET LoadFilter
// Description: Shows how to load an Excel workbook while skipping hidden worksheets by implementing a custom LoadFilter that overrides StartSheet, assigning it to LoadOptions, and saving a new file that contains only the visible sheets.
// Keywords: Aspose.Cells LoadFilter | exclude hidden worksheets | load workbook visible sheets only | C# Aspose.Cells custom LoadFilter | skip hidden sheets Aspose.Cells | LoadOptions visible worksheets | filter worksheets on load .NET | Aspose.Cells performance hidden sheets | Excel hidden sheet loading Aspose
// Common Searches: Aspose.Cells load workbook without hidden sheets | How to ignore hidden worksheets using LoadFilter | C# load only visible worksheets Aspose.Cells | Custom LoadFilter example Aspose.Cells .NET | Skip very hidden sheets Aspose.Cells
// Developer Intent: Load an Excel file and automatically omit hidden worksheets.
// Use Cases: Process large workbooks faster by loading only visible sheets. | Create a distribution copy that contains just the visible worksheets. | Generate a list of visible worksheet names for reporting purposes.
// AI Prompts: Write C# code with Aspose.Cells that loads an Excel workbook while ignoring hidden worksheets by using a custom LoadFilter. | Provide an example of configuring LoadOptions with a LoadFilter to load only visible sheets and then save the filtered workbook. | Explain how to extend the custom LoadFilter to also exclude very‑hidden sheets (IsVeryHidden) in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to load an Excel workbook while skipping hidden worksheets by implementing a custom LoadFilter that overrides StartSheet, assigning it to LoadOptions, and saving a new file that contains only the visible sheets.
    public class ExcludeHiddenSheetsDemo
    {
        // Custom LoadFilter that loads only visible worksheets
        private class VisibleSheetLoadFilter : LoadFilter
        {
            public override void StartSheet(Worksheet sheet)
            {
                // Load data for the sheet only if it is visible
                if (sheet.IsVisible)
                {
                    base.StartSheet(sheet);
                }
            }
        }

        public static void Run()
        {
            // Path to the workbook to be loaded
            string inputFile = "InputWorkbook.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            try
            {
                // Set up LoadOptions with the custom filter
                LoadOptions options = new LoadOptions
                {
                    LoadFilter = new VisibleSheetLoadFilter()
                };

                // Load the workbook; hidden worksheets will be ignored
                Workbook workbook = new Workbook(inputFile, options);

                // List the worksheets that were loaded (should be only visible ones)
                Console.WriteLine("Loaded worksheets:");
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    Console.WriteLine($"- {ws.Name} (Visible = {ws.IsVisible})");
                }

                // Save the filtered workbook
                string outputFile = "FilteredWorkbook.xlsx";
                workbook.Save(outputFile);
                Console.WriteLine($"Filtered workbook saved as: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExcludeHiddenSheetsDemo.Run();
        }
    }
}
