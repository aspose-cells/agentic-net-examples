// Title: Save an Aspose.Cells Workbook as HTML and reload it with LoadOptions to confirm round‑trip data fidelity in C#
// AI Prompts: Export a Workbook to HTML using Aspose.Cells, then load the generated HTML file with LoadOptions and programmatically compare worksheet counts and key cell values. | Write a C# console application that saves a workbook as HTML, reopens it via LoadFormat.Html, and asserts that original data such as header rows remain unchanged.
// Common Searches: Aspose.Cells C# round‑trip workbook to HTML and back using LoadOptions | how to verify data integrity after saving an Excel workbook as HTML with Aspose.Cells | load HTML exported by Aspose.Cells into a new Workbook and compare cell values | C# example for saving a workbook as HTML then reloading with LoadFormat.Html | Aspose.Cells LoadOptions HTML import validation example
// Tags: Aspose.Cells save workbook as HTML | Aspose.Cells import HTML using LoadOptions | HTML round‑trip verification Aspose.Cells | compare worksheet count Aspose.Cells | validate cell value after HTML import Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlRoundTrip
{
    // // Demonstrates creating a simple workbook, exporting it to HTML, reloading the HTML with LoadOptions, and verifying that the worksheet count and the value of cell A1 match between the original and the reloaded workbooks.
    class Program
    {
        static void Main(string[] args)
        {
            // Step 1: Create a simple workbook with some data
            Workbook originalWorkbook = new Workbook();
            Worksheet sheet = originalWorkbook.Worksheets[0];
            sheet.Name = "SampleData";

            // Populate cells with sample values
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Step 2: Save the workbook as HTML
            string htmlPath = "SampleWorkbook.html";
            originalWorkbook.Save(htmlPath, SaveFormat.Html);

            // Step 3: Load the HTML back using LoadOptions
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook loadedWorkbook = new Workbook(htmlPath, loadOptions);

            // Step 4: Verify round‑trip fidelity (basic checks)
            bool isSameSheetCount = originalWorkbook.Worksheets.Count == loadedWorkbook.Worksheets.Count;
            bool isSameFirstCell = originalWorkbook.Worksheets[0].Cells["A1"].StringValue ==
                                   loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;

            Console.WriteLine($"Sheet count match: {isSameSheetCount}");
            Console.WriteLine($"First cell value match: {isSameFirstCell}");

            // Additional verification can be added as needed
        }
    }
}
