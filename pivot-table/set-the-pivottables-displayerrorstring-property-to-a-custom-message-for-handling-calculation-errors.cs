using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableDisplayErrorStringDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                Cells cells = sheet.Cells;
                cells["A1"].Value = "Product";
                cells["B1"].Value = "Sales";
                cells["A2"].Value = "Bike";
                cells["B2"].Value = 1000;
                cells["A3"].Value = "Car";
                cells["B3"].Value = 2000;
                cells["A4"].Value = "Bike";
                cells["B4"].Value = 1500;
                cells["A5"].Value = "Car";
                cells["B5"].Value = 2500;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table (optional, but shows typical usage)
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field

                // Enable custom error string display and set the custom message
                pivotTable.DisplayErrorString = true;
                pivotTable.ErrorString = "Custom Calculation Error";

                // Recalculate the pivot table to apply the settings
                pivotTable.CalculateData();

                // Output the current settings to the console (for verification)
                Console.WriteLine("DisplayErrorString: " + pivotTable.DisplayErrorString);
                Console.WriteLine("ErrorString: " + pivotTable.ErrorString);

                // Ensure the output directory exists
                string outputPath = "PivotTableDisplayErrorStringDemo_out.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTableDisplayErrorStringDemo.Run();
        }
    }
}