// Title: SetFormula to reference an external workbook cell in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create or locate an external workbook, register it as an external link, assign a formula that points to a cell in that workbook using the Formula property (SetFormula), recalculate the workbook, and save the result. The example works with Aspose.Cells for .NET and handles optional custom paths.
// Keywords: Aspose.Cells SetFormula external workbook | C# external link Excel formula | Aspose.Cells reference another file | Add external link Aspose.Cells | Calculate external formulas Aspose.Cells | Excel external workbook formula .NET | Aspose.Cells external workbook example
// Common Searches: Aspose.Cells set formula to another workbook | How to add external link in Aspose.Cells C# | Reference cell in external Excel file using Aspose.Cells | Calculate formulas that pull data from another workbook Aspose.Cells | Save workbook after linking external file Aspose.Cells
// Developer Intent: Insert a formula that pulls a value from a specific cell in an external Excel workbook and ensure the link is registered for correct calculation.
// Use Cases: Create a summary report that always reflects the latest figures from a source workbook. | Consolidate departmental financial data into a master workbook without manual copy‑paste. | Build a dashboard that pulls real‑time metrics from multiple external Excel files.
// AI Prompts: Show C# code using Aspose.Cells to add an external link, set a formula referencing that link, calculate, and save the workbook. | Explain how to register an external workbook in Aspose.Cells and use SetFormula to reference a cell in it. | Provide guidance on handling relative and absolute paths for external workbook formulas with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExternalFormulaDemo
{
    // Demonstrates how to create or locate an external workbook, register it as an external link, assign a formula that points to a cell in that workbook using the Formula property (SetFormula), recalculate the workbook, and save the result. The example works with Aspose.Cells for .NET and handles optional custom paths.
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
                    var extWb = new Workbook();
                    extWb.Worksheets[0].Name = "Sheet1";
                    extWb.Worksheets[0].Cells["A1"].PutValue("External Value");
                    extWb.Save(externalFileName);
                }

                // Create the main workbook
                var mainWb = new Workbook();
                var sheet = mainWb.Worksheets[0];

                // Register the external link to the external workbook and its sheet
                string[] externalSheets = { "Sheet1" };
                int linkIndex = sheet.Workbook.Worksheets.ExternalLinks.Add(externalFileName, externalSheets);
                ExternalLink externalLink = sheet.Workbook.Worksheets.ExternalLinks[linkIndex];

                // Optionally set the data source path if the external file is located elsewhere
                // externalLink.DataSource = @"C:\Path\To\ExternalWorkbook.xlsx";

                // Set a formula that references cell A1 of the external workbook's Sheet1
                sheet.Cells["A1"].Formula = "='[ExternalWorkbook.xlsx]Sheet1'!A1";

                // Calculate formulas (will pull data from the external workbook)
                mainWb.CalculateFormula();

                // Save the main workbook
                mainWb.Save("MainWorkbook_WithExternalFormula.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
