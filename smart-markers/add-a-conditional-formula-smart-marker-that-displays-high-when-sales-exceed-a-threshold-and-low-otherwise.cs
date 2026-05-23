using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace ConditionalSmartMarkerDemoApp
{
    // Demonstrates conditional smart markers using Aspose.Cells
    public class ConditionalSmartMarkerDemo
    {
        // Simple data class for the smart marker data source
        public class SalesInfo
        {
            public double Sales { get; set; }
            public double Threshold { get; set; }
        }

        public static void Run()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Add headers
                sheet.Cells["A1"].PutValue("Sales");
                sheet.Cells["B1"].PutValue("Result");

                // 3. Insert a smart marker formula.
                // The formula uses IF to display "High" when Sales > Threshold, otherwise "Low".
                // Smart markers are referenced with the '&=' prefix.
                sheet.Cells["B2"].PutValue("=IF(&=$Data.Sales>&=$Data.Threshold,\"High\",\"Low\")");

                // 4. Prepare the data source
                SalesInfo data = new SalesInfo
                {
                    Sales = 150,      // Example sales value
                    Threshold = 100   // Example threshold
                };

                // 5. Set up WorkbookDesigner, assign the workbook and data source, then process
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource("Data", data);   // The name "Data" matches the "$Data" prefix in the formula
                designer.Process();                     // Process all smart markers in the workbook

                // 6. Save the result
                string outputPath = "ConditionalSmartMarkerResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main()
        {
            ConditionalSmartMarkerDemo.Run();
        }
    }
}