// Title: Select a worksheet by its name and obtain the Cells collection with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get cells from a worksheet named 'Data' | How to access a specific worksheet's cells using workbook.Worksheets indexer | C# code to select worksheet by name and write to cell A1 with Aspose.Cells | Retrieve Cells object from named worksheet in Aspose.Cells for .NET | Example of selecting a worksheet by name and saving workbook in Aspose.Cells
// Tags: worksheet name selection Aspose.Cells | cells collection retrieval C# | write value to cell A1 Aspose.Cells | Aspose.Cells workbook save example | access named worksheet cells .NET

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a Workbook, adds a worksheet named "Data", selects that worksheet via workbook.Worksheets["Data"], obtains its Cells collection, writes "Hello Aspose.Cells" to cell A1, and saves the file as Output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // creates a new workbook

                // Add a worksheet with a specific name if it does not already exist
                int sheetIndex = workbook.Worksheets.Add();
                Worksheet sheet = workbook.Worksheets[sheetIndex];
                sheet.Name = "Data";

                // Select the worksheet by its name
                Worksheet selectedSheet = workbook.Worksheets["Data"];

                // Obtain the Cells collection of the selected worksheet for further operations
                Cells cells = selectedSheet.Cells;

                // Example usage: write a value to cell A1
                cells["A1"].PutValue("Hello Aspose.Cells");

                // Optionally, save the workbook to verify the result
                string outputPath = "Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
