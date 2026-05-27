using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class HideColumnsShowFormulasAndExport
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate header row
                worksheet.Cells["A1"].PutValue("ID");
                worksheet.Cells["B1"].PutValue("Name");
                worksheet.Cells["C1"].PutValue("Secret");      // This column will be hidden
                worksheet.Cells["D1"].PutValue("Amount");
                worksheet.Cells["E1"].PutValue("Hidden");      // This column will be hidden

                // Populate some data rows
                worksheet.Cells["A2"].PutValue(1);
                worksheet.Cells["B2"].PutValue("Alice");
                worksheet.Cells["C2"].PutValue("TopSecret");
                worksheet.Cells["D2"].PutValue(100);
                worksheet.Cells["E2"].PutValue("HideMe");

                worksheet.Cells["A3"].PutValue(2);
                worksheet.Cells["B3"].PutValue("Bob");
                worksheet.Cells["C3"].PutValue("Classified");
                worksheet.Cells["D3"].PutValue(200);
                worksheet.Cells["E3"].PutValue("HideMeToo");

                // Add formulas to demonstrate formula display
                worksheet.Cells["F2"].Formula = "=D2*2";
                worksheet.Cells["F3"].Formula = "=D3*2";

                // Define header names that should be hidden
                HashSet<string> headersToHide = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "Secret",
                    "Hidden"
                };

                // Hide matching columns based on header text
                int maxColumn = worksheet.Cells.MaxColumn; // inclusive index of last used column
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell headerCell = worksheet.Cells[0, col];
                    if (headerCell != null && headerCell.Type == CellValueType.IsString)
                    {
                        string headerText = headerCell.StringValue;
                        if (headersToHide.Contains(headerText))
                        {
                            worksheet.Cells.HideColumn(col);
                        }
                    }
                }

                // Show formulas instead of calculated results
                worksheet.ShowFormulas = true;

                // Export visible data to a DataTable (optional demonstration)
                ExportTableOptions exportOptions = new ExportTableOptions
                {
                    PlotVisibleColumns = true,
                    PlotVisibleRows = true,
                    PlotVisibleCells = true,
                    ExportColumnName = true
                };
                DataTable visibleData = worksheet.Cells.ExportDataTable(
                    0,
                    0,
                    worksheet.Cells.MaxRow + 1,
                    worksheet.Cells.MaxColumn + 1,
                    exportOptions);

                Console.WriteLine("Exported DataTable Columns:");
                foreach (DataColumn col in visibleData.Columns)
                {
                    Console.Write(col.ColumnName + "\t");
                }
                Console.WriteLine();

                // Save the workbook to XLSX format
                string outputPath = "HiddenColumns_ShowFormulas.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            HideColumnsShowFormulasAndExport.Run();
        }
    }
}