// Title: Create a new Aspose.Cells Workbook and add a "PivotData" worksheet for pivot tables in C#
// Common Searches: how to add a worksheet named PivotData to a new Aspose.Cells workbook in C# | C# Aspose.Cells example to create a workbook and retrieve the index of a newly added worksheet | save an Aspose.Cells workbook to a specific folder after creating the output directory | prepare a worksheet for pivot table data using Aspose.Cells in a .NET application
// Tags: Aspose.Cells create workbook C# | Aspose.Cells add pivot data worksheet | Aspose.Cells get worksheet index | Aspose.Cells save workbook as XLSX | Aspose.Cells ensure output directory exists

using Aspose.Cells;
using System;
using System.IO;

// The sample demonstrates how to instantiate a new Aspose.Cells Workbook, add a worksheet called "PivotData" to hold pivot‑table source data, obtain the worksheet's index, ensure the target directory exists, and save the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Add a worksheet that will hold the data for the pivot table
            Worksheet pivotWorksheet = workbook.Worksheets.Add("PivotData");

            // (Optional) Get the index of the newly added worksheet
            int sheetIndex = workbook.Worksheets.IndexOf(pivotWorksheet);

            // Save the workbook to verify creation
            string outputPath = "PivotWorkbook.xlsx";

            // Ensure the directory exists before saving
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'. Worksheet index: {sheetIndex}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
