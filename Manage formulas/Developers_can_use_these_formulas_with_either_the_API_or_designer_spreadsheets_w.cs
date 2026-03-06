using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaDemo
{
    class Program
    {
        static void Main()
        {
            string templatePath = "TemplateWithFormulas.xlsx";

            // Create a simple template workbook if it does not exist
            if (!File.Exists(templatePath))
            {
                var tempWb = new Workbook();
                var ws = tempWb.Worksheets[0];
                ws.Name = "Data";

                // Header
                ws.Cells["A1"].PutValue("Opportunity_Name");
                ws.Cells["B1"].PutValue("Opportunity_Amount");
                ws.Cells["C1"].PutValue("Total");

                // Sample data rows (will be replaced by smart markers)
                ws.Cells["A2"].PutValue("&=Master.Opportunity_Name");
                ws.Cells["B2"].PutValue("&=Master.Opportunity_Amount");
                // Formula that sums numeric part of amount (for demo)
                ws.Cells["C2"].Formula = "SUM(B2)";

                tempWb.Save(templatePath, SaveFormat.Xlsx);
            }

            // Load the workbook with formula parsing enabled
            var loadOptions = new LoadOptions { ParsingFormulaOnOpen = true };
            var workbook = new Workbook(templatePath, loadOptions);

            // Prepare data source
            var dt = new DataTable("Master");
            dt.Columns.Add("Opportunity_Name");
            dt.Columns.Add("Opportunity_Amount");
            var dr = dt.NewRow();
            dr["Opportunity_Name"] = "Test Deal";
            dr["Opportunity_Amount"] = 2500.0; // numeric value for proper calculation
            dt.Rows.Add(dr);

            // Bind data source and process smart markers
            var designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.CalculateFormula = true;
            designer.Process();

            // Ensure any remaining formulas are calculated
            workbook.CalculateFormula();

            // Save result
            string outputPath = "ResultWithCalculatedFormulas.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Display a sample result
            var sheet = workbook.Worksheets[0];
            Console.WriteLine("Formula in C2 result: " + sheet.Cells["C2"].Value);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
    }
}