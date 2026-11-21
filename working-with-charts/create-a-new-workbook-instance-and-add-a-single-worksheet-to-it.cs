// Title: Create a new Aspose.Cells Workbook, rename the default sheet, and add an extra worksheet using C# (.NET)
// Common Searches: C# Aspose.Cells create workbook and add a new worksheet programmatically | Aspose.Cells .NET example for initializing a workbook and inserting a sheet | How to rename the default worksheet and add another sheet with Aspose.Cells in C# | Saving a newly created workbook to an .xlsx file using Aspose.Cells C# | Aspose.Cells workbook creation with exception handling in .NET
// Tags: Aspose.Cells workbook creation C# | add worksheet Aspose.Cells .NET | rename default sheet Aspose.Cells | save workbook as xlsx Aspose.Cells | exception handling Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample demonstrates how to instantiate a new Aspose.Cells Workbook in C#, rename the automatically added default worksheet to "Sheet1", add a second worksheet named "Sheet2", and save the file as Result.xlsx, including basic exception handling.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (contains a default worksheet)
                Workbook workbook = new Workbook();

                // Rename the default worksheet to a unique name
                Worksheet defaultSheet = workbook.Worksheets[0];
                defaultSheet.Name = "Sheet1";

                // Add an additional worksheet with a different name (optional)
                Worksheet additionalSheet = workbook.Worksheets.Add("Sheet2");

                // Save the workbook to verify the result
                string outputPath = "Result.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
