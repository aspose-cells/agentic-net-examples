// Title: Rename an Excel table (ListObject) to a project‑standard identifier using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells in C# to assign the DisplayName of the first ListObject to "tblSalesData" and save the workbook. | Programmatically change an Excel table's name to follow project naming conventions with the Aspose.Cells .NET API. | Update the identifier of a worksheet table and persist the changes using Aspose.Cells in a C# console application.
// Common Searches: Aspose.Cells C# example for setting ListObject DisplayName | how to change ListObject name in an Excel workbook using Aspose.Cells | set custom table identifier in Excel file with Aspose.Cells API | C# code to rename Excel table to tblSalesData via Aspose.Cells | rename worksheet table programmatically with Aspose.Cells for .NET
// Tags: Aspose.Cells assign table DisplayName | change Excel table name using Aspose.Cells .NET | C# update Excel table naming | Aspose.Cells modify table name in workbook | programmatic Excel table naming with Aspose API

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The sample loads an existing workbook, accesses the first worksheet, ensures a ListObject is present, assigns a new DisplayName (e.g., tblSalesData) to follow naming conventions, creates the output directory if needed, and saves the modified workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables (ListObjects) found in the worksheet.");
                return;
            }

            // Retrieve the first table (ListObject) on the worksheet
            ListObject table = sheet.ListObjects[0];

            // Rename the table using the correct property (DisplayName)
            table.DisplayName = "tblSalesData";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
