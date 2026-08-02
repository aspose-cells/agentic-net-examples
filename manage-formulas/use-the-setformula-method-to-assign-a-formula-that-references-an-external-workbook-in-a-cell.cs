// Title: Set an external workbook reference formula with Aspose.Cells for .NET (C#)
// Description: This example creates (or verifies) an external Excel file, adds it as an ExternalLink to a new workbook, and uses the SetFormula method (via the Formula property) to insert a formula in cell B2 that points to cell A1 on Sheet1 of the external workbook. The workbook is then saved.
// Keywords: Aspose.Cells SetFormula external workbook | C# Aspose.Cells external link | reference another workbook Excel .NET | ExternalLink example Aspose.Cells | link external Excel file C# | GitHub Aspose.Cells external formula sample | global Excel automation
// Common Searches: how to reference a cell in another workbook using Aspose.Cells C# | Aspose.Cells external link and formula tutorial | set formula to external Excel file with Aspose.Cells | C# code for external workbook formula Aspose.Cells
// Developer Intent: Insert a formula that pulls data from a specific cell in an external Excel workbook using Aspose.Cells for .NET.
// Use Cases: Build a consolidated financial dashboard that pulls key figures from departmental workbooks. | Create a master report that references lookup tables stored in separate Excel files. | Maintain a single source of truth for pricing data and link it to multiple analysis workbooks.
// AI Prompts: Write C# code with Aspose.Cells to add an ExternalLink to a workbook and set a formula that references cell A1 of Sheet1 in the linked file. | Explain how to retrieve the index of an added ExternalLink and construct the correct formula string for SetFormula. | Suggest robust error‑handling strategies when the external workbook is missing or the link path is invalid while using SetFormula.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalFormulaDemo
{
    // This example creates (or verifies) an external Excel file, adds it as an ExternalLink to a new workbook, and uses the SetFormula method (via the Formula property) to insert a formula in cell B2 that points to cell A1 on Sheet1 of the external workbook. The workbook is then saved.
    class Program
    {
        static void Main()
        {
            try
            {
                // Ensure the external workbook exists; create a simple one if missing
                string externalFileName = "ExternalWorkbook.xlsx";
                if (!File.Exists(externalFileName))
                {
                    var externalWb = new Workbook();
                    Worksheet extSheet = externalWb.Worksheets[0];
                    extSheet.Name = "Sheet1";
                    extSheet.Cells["A1"].PutValue("External Value");
                    externalWb.Save(externalFileName);
                }

                // Create the main workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add an external link to the main workbook
                string[] externalSheets = new string[] { "Sheet1" };
                int linkIndex = workbook.Worksheets.ExternalLinks.Add(externalFileName, externalSheets);
                ExternalLink externalLink = workbook.Worksheets.ExternalLinks[linkIndex];

                // Optional: set the data source path if needed
                // externalLink.DataSource = Path.GetFullPath(externalFileName);

                // Set a formula that references cell A1 in the external workbook
                sheet.Cells["B2"].Formula = "='[ExternalWorkbook.xlsx]Sheet1'!A1";

                // Save the workbook
                string outputFile = "OutputWithExternalFormula.xlsx";
                workbook.Save(outputFile);

                Console.WriteLine($"Workbook saved successfully as '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
