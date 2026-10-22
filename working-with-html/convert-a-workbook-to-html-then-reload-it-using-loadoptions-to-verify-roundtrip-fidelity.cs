// Title: Convert an Excel workbook to HTML and reload it with LoadOptions to validate round‑trip fidelity using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that saves a Workbook as HTML, then loads the HTML file with LoadOptions(LoadFormat.Html) and checks that sheet names and selected cell values match the original. | Create a reusable C# method that exports a Workbook to HTML, reimports it using Aspose.Cells LoadOptions, and returns a boolean indicating whether the round‑trip conversion preserved data integrity.
// Common Searches: Aspose.Cells C# export workbook to HTML and import it back preserving data | verify round‑trip conversion from Excel to HTML using Aspose.Cells .NET | load HTML file into Aspose.Cells workbook with LoadOptions and compare cell values | check if worksheet names and cell contents remain unchanged after saving as HTML in Aspose.Cells
// Tags: export workbook to html Aspose.Cells | import html workbook with LoadOptions Aspose.Cells | round‑trip data integrity Excel HTML Aspose.Cells | compare worksheet cell values C# Aspose.Cells | loadformat.html Aspose.Cells example

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlRoundTrip
{
    // The program creates a sample workbook, saves it as an HTML file, reloads the file using LoadOptions(LoadFormat.Html), and verifies that the worksheet count, sheet names, and specific cell values are identical to the original, confirming round‑trip fidelity.
    class Program
    {
        static void Main(string[] args)
        {
            // 1. Create a sample workbook with some data
            Workbook originalWorkbook = new Workbook();
            Worksheet sheet = originalWorkbook.Worksheets[0];
            sheet.Name = "SampleData";

            // Populate some cells
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // 2. Save the workbook to HTML format
            string htmlPath = "sample.html";
            originalWorkbook.Save(htmlPath, SaveFormat.Html);

            // 3. Load the HTML back into a new workbook using LoadOptions
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook roundTripWorkbook = new Workbook(htmlPath, loadOptions);

            // 4. Verify round‑trip fidelity (basic checks)
            bool isFidelityOk = true;

            // Check worksheet count
            if (originalWorkbook.Worksheets.Count != roundTripWorkbook.Worksheets.Count)
                isFidelityOk = false;

            // Check sheet names and a few cell values
            for (int i = 0; i < originalWorkbook.Worksheets.Count && isFidelityOk; i++)
            {
                Worksheet origSheet = originalWorkbook.Worksheets[i];
                Worksheet rtSheet = roundTripWorkbook.Worksheets[i];

                if (origSheet.Name != rtSheet.Name)
                {
                    isFidelityOk = false;
                    break;
                }

                // Compare a subset of cells
                string[] cellsToCheck = { "A1", "B1", "A2", "B2", "A3", "B3" };
                foreach (string address in cellsToCheck)
                {
                    var origValue = origSheet.Cells[address].Value?.ToString() ?? string.Empty;
                    var rtValue = rtSheet.Cells[address].Value?.ToString() ?? string.Empty;
                    if (!origValue.Equals(rtValue, StringComparison.Ordinal))
                    {
                        isFidelityOk = false;
                        break;
                    }
                }
            }

            // Output the verification result
            Console.WriteLine(isFidelityOk
                ? "Round‑trip fidelity verified successfully."
                : "Round‑trip fidelity check failed.");
        }
    }
}
