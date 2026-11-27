// Title: How to select a worksheet by name or index and rename it with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get worksheet by its name example | retrieve worksheet using zero‑based index with Aspose.Cells | rename an Excel sheet programmatically using Aspose.Cells .NET | add worksheets until a specific index exists in Aspose.Cells workbook
// Tags: worksheet retrieval by name Aspose.Cells | worksheet retrieval by index Aspose.Cells | rename worksheet Aspose.Cells | add missing worksheets to reach index Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads or creates an Excel workbook, guarantees at least three sheets, obtains a sheet by its name (creating it if missing), obtains another sheet by zero‑based index (adding sheets as needed), renames the named sheet, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a simple workbook if it doesn't.
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Name = "DataSheet";
                // Add extra sheets so index 2 is valid.
                wb.Worksheets.Add();
                wb.Worksheets.Add();
                wb.Save(inputPath);
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Retrieve worksheet by name.
            Worksheet sheetByName = workbook.Worksheets["DataSheet"];
            if (sheetByName == null)
            {
                // Create the sheet if it does not exist.
                int idx = workbook.Worksheets.Add();
                sheetByName = workbook.Worksheets[idx];
                sheetByName.Name = "DataSheet";
            }

            // Retrieve worksheet by zero‑based index (2).
            Worksheet sheetByIndex;
            if (workbook.Worksheets.Count > 2)
            {
                sheetByIndex = workbook.Worksheets[2];
            }
            else
            {
                // Add missing sheets up to the required index.
                while (workbook.Worksheets.Count <= 2)
                {
                    workbook.Worksheets.Add();
                }
                sheetByIndex = workbook.Worksheets[2];
            }

            // Example operation: rename the sheet selected by name.
            sheetByName.Name = "RenamedSheet";

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
