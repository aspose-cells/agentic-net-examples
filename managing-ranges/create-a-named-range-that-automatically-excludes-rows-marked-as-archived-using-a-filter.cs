// Title: How to create a dynamic named range in Aspose.Cells (C#) that automatically skips rows where Status = "Archived"
// AI Prompts: Generate C# code with Aspose.Cells that defines a workbook‑level named range using an OFFSET/COUNTIF formula to include only rows whose Status column is not "Archived". | Show how to add a named range called ActiveData that updates automatically when rows are added or removed, excluding archived entries, via the Aspose.Cells API.
// Common Searches: Aspose.Cells C# create named range that excludes rows with value 'Archived' in a column | dynamic Excel named range using OFFSET and COUNTIF with Aspose.Cells .NET | filter out archived rows when defining a named range in an Aspose.Cells workbook | C# Aspose.Cells example for conditional named range based on status column
// Tags: Aspose.Cells OFFSET formula | exclude archived rows Aspose.Cells | COUNTIF based named range C# | Excel workbook conditional named range | filter rows in named range .NET

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, fills it with sample data containing ID, Name, and Status columns, then defines a dynamic named range called "ActiveData" using an OFFSET formula that subtracts rows marked as "Archived" via COUNTIF, and finally saves the file as an .xlsx workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Add header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Status");

            // Sample data (including rows marked as "Archived")
            object[,] sampleData = new object[,]
            {
                { 1, "Alice",   "Active"   },
                { 2, "Bob",     "Archived" },
                { 3, "Charlie", "Active"   },
                { 4, "Diana",   "Archived" },
                { 5, "Eve",     "Active"   }
            };

            // Import the data starting at cell A2 (row index 1, column index 0)
            // Aspose.Cells provides ImportArray2D, but to ensure compatibility we fill manually.
            int rows = sampleData.GetLength(0);
            int cols = sampleData.GetLength(1);
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    sheet.Cells[r + 1, c].PutValue(sampleData[r, c]);
                }
            }

            // Create a dynamic named range "ActiveData" that excludes rows where Status = "Archived"
            string formula = "=OFFSET(Data!$A$2,0,0,COUNTA(Data!$A:$A)-COUNTIF(Data!$C:$C,\"Archived\"),3)";

            // Add the named range to the workbook
            int nameIndex = workbook.Worksheets.Names.Add("ActiveData");
            workbook.Worksheets.Names[nameIndex].RefersTo = formula;

            // Define output file path
            string outputPath = "NamedRange_ExcludeArchived.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
