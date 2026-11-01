// Title: Create a dynamic named range that automatically expands when adding columns to an Aspose.Cells ListObject (C#)
// Common Searches: Aspose.Cells C# create named range that expands with added columns | use structured reference MyTable[#All] in Aspose.Cells for a dynamic range | C# Aspose.Cells ListObject auto expanding named range example | how to set RefersTo = "=MyTable[#All]" with Aspose.Cells .NET | auto expanding Excel named range using Aspose.Cells API in C#
// Tags: Aspose.Cells structured reference range | C# ListObject auto expanding range | Excel table #All reference .NET | programmatic named range Aspose.Cells | auto expanding Excel range using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample creates a new workbook, adds a ListObject named MyTable covering cells A1:B2, defines a named range MyDynamicRange that refers to =MyTable[#All] so it automatically expands when new columns are added, and saves the file as DynamicNamedRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some initial data for the table
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["B2"].PutValue(20);

            // Define the range that will become the table (A1:B2)
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int totalRows = 2;     // includes header row
            int totalColumns = 2;  // A and B columns

            // Add a ListObject (Excel table) to the worksheet.
            // The last parameter 'hasHeaders' is set to true because the first row contains headers.
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Use DisplayName instead of Name (Aspose.Cells API)
            table.DisplayName = "MyTable";
            table.ShowHeaderRow = true;
            table.ShowTableStyleFirstColumn = false;
            table.ShowTableStyleLastColumn = false;

            // Create a named range that refers to the entire table using a structured reference.
            // The structured reference (=MyTable[#All]) automatically expands when new columns are added.
            int nameIndex = workbook.Worksheets.Names.Add("MyDynamicRange");
            Name dynamicRange = workbook.Worksheets.Names[nameIndex];
            dynamicRange.RefersTo = "=MyTable[#All]";

            // Define output file path
            string outputPath = "DynamicNamedRange.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

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
