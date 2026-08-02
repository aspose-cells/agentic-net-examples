using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerCurrency
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Add a worksheet that will contain the smart marker template
                Worksheet templateSheet = workbook.Worksheets[0];
                templateSheet.Name = "Template";

                // Place a smart marker that concatenates a literal currency symbol with the numeric field "Amount"
                // Using "$" directly instead of a variable to avoid missing SetVariableValue API
                templateSheet.Cells["A1"].PutValue("$&Amount");

                // Prepare data source: a DataTable with a monetary column
                DataTable dt = new DataTable("Data");
                dt.Columns.Add("Amount", typeof(double));
                dt.Rows.Add(1234.56); // sample monetary value

                // Create a WorkbookDesigner to work with smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Assign the data source to the designer
                designer.SetDataSource(dt);

                // Process the smart markers (they will be replaced with actual values)
                designer.Process();

                // Apply a custom number format to the result cell to keep numeric formatting
                Style style = workbook.CreateStyle();
                style.Custom = "#,##0.00"; // standard monetary format without symbol
                StyleFlag flag = new StyleFlag { NumberFormat = true };

                // Create a range covering cell A1 (row 0, column 0)
                Aspose.Cells.Range resultRange = templateSheet.Cells.CreateRange(0, 0, 1, 1);
                resultRange.ApplyStyle(style, flag);

                // Ensure the output directory exists
                string outputPath = "SmartMarkerCurrencyDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (lifecycle rule: save)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}