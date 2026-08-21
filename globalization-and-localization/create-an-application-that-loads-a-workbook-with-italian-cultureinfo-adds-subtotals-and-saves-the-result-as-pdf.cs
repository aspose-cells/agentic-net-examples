// Title: C# – Load Excel with Italian CultureInfo, add subtotals, export to PDF using Aspose.Cells
// Description: A concise Aspose.Cells for .NET demo that loads an XLSX file with Italian (it‑IT) CultureInfo, inserts a subtotal grouping sales by category, and saves the result as a PDF document.
// Keywords: Aspose.Cells C# | .NET Excel to PDF | Italian CultureInfo it-IT | LoadOptions CultureInfo | Cells.Subtotal method | Excel subtotal example | localized Excel report | PDF export Aspose.Cells | group by category Excel | sum function Excel C#
// Common Searches: Aspose.Cells load workbook with Italian locale | C# add subtotal to Excel sheet using Aspose.Cells | Export Excel with subtotals to PDF .NET | How to set CultureInfo it-IT in Aspose.Cells LoadOptions | Subtotal function example Aspose.Cells C#
// Developer Intent: Load an existing XLSX file using Italian locale settings, apply a subtotal that sums sales per category, and generate a PDF version of the workbook.
// Use Cases: Create printable Italian sales reports that automatically calculate category totals. | Automate generation of inventory sheets for the Italian market with summed quantities before PDF distribution. | Produce localized financial summaries in PDF format by applying subtotals to raw Excel data.
// AI Prompts: Provide C# code that loads an XLSX with LoadOptions.CultureInfo set to "it-IT", adds a subtotal on the sales column grouped by category, and saves the workbook as a PDF using Aspose.Cells. | Show an Aspose.Cells example for applying Italian CultureInfo, using Cells.Subtotal to summarize data, and exporting the result to PDF. | Explain how to configure LoadOptions for Italian locale and use the Subtotal method to generate a PDF report in .NET.

using System;
using System.Globalization;
using Aspose.Cells;

namespace SubtotalItalianPdfDemo
{
    // A concise Aspose.Cells for .NET demo that loads an XLSX file with Italian (it‑IT) CultureInfo, inserts a subtotal grouping sales by category, and saves the result as a PDF document.
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a sample workbook with data and save it as an XLSX file.
            // -----------------------------------------------------------------
            Workbook tempWorkbook = new Workbook();                         // create workbook
            Worksheet sheet = tempWorkbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Categoria");   // Category (Italian)
            cells["B1"].PutValue("Prodotto");   // Product
            cells["C1"].PutValue("Vendite");    // Sales

            // Sample data rows
            object[,] data = new object[,]
            {
                { "Nord", "Widget", 5000 },
                { "Nord", "Gadget", 3000 },
                { "Sud",  "Widget", 6000 },
                { "Sud",  "Gadget", 4000 },
                { "Ovest","Widget", 4500 }
            };

            for (int r = 0; r < data.GetLength(0); r++)
                for (int c = 0; c < data.GetLength(1); c++)
                    cells[r + 1, c].PutValue(data[r, c]);

            // Save the temporary Excel file (will be loaded later with Italian culture)
            string tempFilePath = "sample.xlsx";
            tempWorkbook.Save(tempFilePath, SaveFormat.Xlsx);

            // ---------------------------------------------------------------
            // 2. Load the workbook using Italian CultureInfo (it-IT)
            // ---------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.CultureInfo = new CultureInfo("it-IT");   // set Italian culture for loading
            Workbook workbook = new Workbook(tempFilePath, loadOptions); // load with options

            // ---------------------------------------------------------------
            // 3. Add subtotals to the loaded workbook
            // ---------------------------------------------------------------
            Worksheet ws = workbook.Worksheets[0];
            Cells wsCells = ws.Cells;

            // Define the range that contains the data (A1:C6)
            CellArea area = CellArea.CreateCellArea(0, 0, 5, 2); // rows 0-5, columns 0-2

            // Group by the first column (Categoria), sum the sales column (index 2)
            wsCells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 2 });

            // ---------------------------------------------------------------
            // 4. Save the result as PDF
            // ---------------------------------------------------------------
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine("Workbook loaded with Italian culture, subtotal added, and saved as PDF.");
        }
    }
}
