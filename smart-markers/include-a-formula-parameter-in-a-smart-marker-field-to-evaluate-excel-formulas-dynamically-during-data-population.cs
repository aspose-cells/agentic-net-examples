// Title: How to use a Formula parameter in Aspose.Cells smart markers to calculate totals dynamically in C#
// AI Prompts: Generate C# code that creates an Excel template with smart markers, adds a smart marker using the Formula parameter (e.g., &=SUM(Price)), binds a DataTable, enables WorkbookDesigner.CalculateFormula, processes the markers, and saves the workbook. | Write a C# snippet that reads the evaluated value of a cell containing a formula after smart marker processing with WorkbookDesigner. | Adapt an existing Aspose.Cells smart‑marker example to insert a custom Excel function via the Formula parameter and output the computed result.
// Common Searches: Aspose.Cells C# smart marker with formula parameter example | How to calculate a SUM column using smart markers in Aspose.Cells | Enable WorkbookDesigner.CalculateFormula to evaluate formulas after smart marker processing | Insert Excel formula via smart marker &=SUM in C# code | Populate Excel from DataTable and compute totals with Aspose.Cells smart markers
// Tags: smart marker formula parameter Aspose.Cells | WorkbookDesigner calculate formula C# | populate Excel using DataTable smart markers | dynamic SUM formula insertion Aspose.Cells | C# evaluate Excel formulas after smart marker processing

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerFormulaDemo
{
    // Shows how to build an in‑memory Excel template with smart markers, embed a Formula parameter (e.g., &=SUM(Price)) to insert a SUM formula, bind a DataTable as the data source, enable WorkbookDesigner.CalculateFormula for automatic evaluation, process the markers, and save the workbook while displaying the resulting formula and its calculated value.
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a template workbook in memory and define smart markers
            // ------------------------------------------------------------
            Workbook templateWb = new Workbook();
            Worksheet tmplSheet = templateWb.Worksheets[0];
            Cells tmplCells = tmplSheet.Cells;

            // Header row
            tmplCells["A1"].PutValue("Product");
            tmplCells["B1"].PutValue("Price");
            tmplCells["C1"].PutValue("Total Price");

            // Data rows – smart markers that will be replaced by data source values
            tmplCells["A2"].PutValue("&=$Product");   // product name
            tmplCells["B2"].PutValue("&=$Price");    // price value

            // Smart marker with a Formula parameter.
            // After processing, this cell will contain a formula that sums the entire Price column.
            tmplCells["C2"].PutValue("&=SUM(Price)");

            // Save the template to a memory stream (simulating a file)
            using (MemoryStream tmplStream = new MemoryStream())
            {
                templateWb.Save(tmplStream, SaveFormat.Xlsx);
                tmplStream.Position = 0; // Reset stream position for reading

                // ------------------------------------------------------------
                // 2. Load the template into WorkbookDesigner
                // ------------------------------------------------------------
                WorkbookDesigner designer = new WorkbookDesigner();
                designer.Workbook = new Workbook(tmplStream);

                // ------------------------------------------------------------
                // 3. Prepare the data source (a DataTable)
                // ------------------------------------------------------------
                DataTable dt = new DataTable("Products");
                dt.Columns.Add("Product", typeof(string));
                dt.Columns.Add("Price", typeof(double));

                dt.Rows.Add("Apple", 1.20);
                dt.Rows.Add("Banana", 0.80);
                dt.Rows.Add("Cherry", 2.50);
                dt.Rows.Add("Date", 3.00);

                // ------------------------------------------------------------
                // 4. Bind the data source and enable formula calculation
                // ------------------------------------------------------------
                designer.SetDataSource(dt);
                designer.CalculateFormula = true; // Calculate formulas after smart marker processing

                // ------------------------------------------------------------
                // 5. Process the smart markers
                // ------------------------------------------------------------
                designer.Process();

                // ------------------------------------------------------------
                // 6. Save the populated workbook to a file
                // ------------------------------------------------------------
                string outputPath = "SmartMarker_With_Formula.xlsx";
                designer.Workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
                // Verify that the formula in C2 has been evaluated
                Console.WriteLine($"C2 Formula: {designer.Workbook.Worksheets[0].Cells["C2"].Formula}");
                Console.WriteLine($"C2 Value (calculated): {designer.Workbook.Worksheets[0].Cells["C2"].Value}");
            }
        }
    }
}
