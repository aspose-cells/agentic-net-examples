// Title: Import a POCO collection into an Excel worksheet while preserving merged cells using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, merges a range, and imports a List<SampleData> with ImportTableOptions.CheckMergedCells enabled. | Show how to set ImportTableOptions to insert rows so the imported data does not overwrite existing merged cells. | Provide the steps to save the workbook as an .xlsx file after the import and verify that the merged range remains intact.
// Common Searches: Aspose.Cells C# import custom objects without breaking merged cells | How to enable CheckMergedCells when using ImportCustomObjects in .NET | Insert rows while importing a collection into an existing worksheet with Aspose.Cells | Preserve merged cell range D4:D5 during data import using Aspose.Cells for .NET | Import POCO list into Excel starting at A1 with Aspose.Cells
// Tags: ImportCustomObjects CheckMergedCells option | Aspose.Cells merged cells preservation | C# import POCO list into worksheet | InsertRows during data import Aspose.Cells | Save workbook as Xlsx format

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportMergedDemo
{
    // Simple POCO to represent data rows
    // The example creates a new workbook, merges cells D4:D5, builds a List<SampleData>, configures ImportTableOptions with CheckMergedCells and InsertRows enabled, imports the collection starting at A1, and saves the file as ImportWithMergedCells.xlsx.
    public class SampleData
    {
        public int Id { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Create a merged range that will be overlapped by imported data
                // Merge cells D4:D5 (zero‑based indices: row 3, column 3, 2 rows, 1 column)
                cells.Merge(3, 3, 2, 1);
                cells[3, 3].PutValue("Pre‑merged");

                // Prepare sample data to import; it has two columns and three rows
                List<SampleData> data = new List<SampleData>
                {
                    new SampleData { Id = 1, Description = "First" },
                    new SampleData { Id = 2, Description = "Second" },
                    new SampleData { Id = 3, Description = "Third" }
                };

                // Configure import options: enable merged‑cell checking
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    IsFieldNameShown = false,   // Do not import property names as header
                    InsertRows = true,          // Insert rows instead of overwriting existing ones
                    CheckMergedCells = true     // Respect merged cells during import
                };

                // Import the custom objects starting at cell A1 (row 0, column 0)
                // The data will occupy columns A and B; rows will be inserted below the merged range
                cells.ImportCustomObjects((ICollection)data, 0, 0, importOptions);

                // Save the workbook to verify the result
                workbook.Save("ImportWithMergedCells.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
