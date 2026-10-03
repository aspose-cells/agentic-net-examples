// Title: How to apply a US‑locale custom currency number format to a smart marker using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that inserts a smart marker, creates a style with the format string "[$-en-US]$#,##0.00", applies the style to the marker cell, and processes the workbook with WorkbookDesigner. | Demonstrate binding a DataSet to a smart marker and formatting the resulting cell as a monetary value in the generated workbook. | Provide a complete example that saves the workbook to an .xlsx file, creates the output directory if needed, and includes basic error handling.
// Common Searches: aspnet apply US dollar number format to smart marker Aspose.Cells | c# Aspose.Cells WorkbookDesigner format cell as currency with locale | how to use format string [$-en-US] with smart markers in Excel using Aspose | set locale specific currency style for smart marker data source Aspose.Cells C# example
// Tags: US currency formatting smart marker | WorkbookDesigner custom number style | apply custom number style to cell Aspose.Cells | export Excel with currency formatting C# | smart marker data binding Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerCurrencyExample
{
    // The example creates a workbook, places a smart marker "&[Amount]" in cell A1, defines a US‑locale custom currency format "[$-en-US]$#,##0.00", applies this style to the marker cell, binds a DataSet containing an Amount column (1234.56) to the smart marker, processes it with WorkbookDesigner, and saves the result as SmartMarkerCurrency.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Insert a smart marker that will be replaced by a currency value
                sheet.Cells["A1"].PutValue("&[Amount]");

                // Define a custom number format for currency (US locale)
                Style currencyStyle = workbook.CreateStyle();
                currencyStyle.Custom = "[$-en-US]$#,##0.00";

                // Apply the custom style to the cell containing the smart marker
                sheet.Cells["A1"].SetStyle(currencyStyle);

                // Prepare the data source for the smart marker using a DataSet
                DataTable table = new DataTable();
                table.Columns.Add("Amount", typeof(double));
                table.Rows.Add(1234.56);

                DataSet dataSet = new DataSet();
                dataSet.Tables.Add(table);

                // Process the smart marker using WorkbookDesigner
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dataSet);
                designer.Process();

                // Determine output path and ensure the directory exists
                string outputFile = "SmartMarkerCurrency.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));

                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the resulting workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while generating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
