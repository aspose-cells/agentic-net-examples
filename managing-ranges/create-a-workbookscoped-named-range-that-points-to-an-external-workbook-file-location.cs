// Title: Create a workbook‑scoped named range that references an external Excel file with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add workbook scoped named range that points to another workbook file | how to set RefersTo property for an external workbook reference using Aspose.Cells | example of building external reference string for a named range in Aspose.Cells .NET | C# code to create a named range referencing external sheet A1:B10 with Aspose.Cells | save a workbook that contains an external named range using Aspose.Cells for .NET
// Tags: workbook scoped named range external workbook Aspose.Cells | RefersTo property external reference C# | Aspose.Cells create external named range Excel | add named range with external file path Aspose.Cells | external workbook reference string format Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates an in‑memory workbook, constructs an external reference string in the Excel format "'[full path]sheet name'!range", adds a workbook‑scoped named range named "ExternalData" via Worksheets.Names.Add, assigns the external reference to its RefersTo property, and saves the workbook as WorkbookWithExternalNamedRange.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (in-memory)
                Workbook workbook = new Workbook();

                // Define the external workbook reference.
                // Example: external file located at C:\Data\ExternalWorkbook.xlsx, sheet named "Sheet1",
                // range A1:B10.
                string externalFilePath = @"C:\Data\ExternalWorkbook.xlsx";
                string externalSheetName = "Sheet1";
                string externalRange = "$A$1:$B$10";

                // Ensure the external file exists to avoid FileNotFoundException.
                if (!File.Exists(externalFilePath))
                {
                    Console.WriteLine($"External file not found: {externalFilePath}");
                    return;
                }

                // Build the external reference string in the format required by Excel:
                // '[full path]sheet name'!range
                string externalReference = string.Format("'[{0}]{1}'!{2}",
                    externalFilePath, externalSheetName, externalRange);

                // Add a workbook‑scoped named range that points to the external location.
                // The Add method returns the index of the newly added name.
                int nameIndex = workbook.Worksheets.Names.Add("ExternalData");
                Name externalName = workbook.Worksheets.Names[nameIndex];
                externalName.RefersTo = externalReference;

                // Save the workbook to a file.
                string outputPath = "WorkbookWithExternalNamedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
