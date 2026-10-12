// Title: Rename the first Excel table to a project‑specific identifier using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a workbook, finds the first ListObject, and sets its DisplayName to tbl_{WorksheetName}_1 using Aspose.Cells. | Create a script that checks for tables in a worksheet and renames them according to a naming convention (tbl_{SheetName}_{Index}) with Aspose.Cells .NET. | Write a C# routine that validates the input file, renames an Excel table's DisplayName to follow project conventions, and saves the workbook.
// Common Searches: how to change Excel table name programmatically with Aspose.Cells C# | Aspose.Cells set ListObject DisplayName based on worksheet name | C# rename Excel table to follow naming convention using Aspose.Cells | example of renaming Excel table in .xlsx file with Aspose.Cells .NET
// Tags: Aspose.Cells rename ListObject DisplayName | C# set Excel table name with Aspose.Cells | project naming convention for Excel tables .NET | update Excel table identifier using Aspose.Cells | load and save workbook after table rename Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Loads input.xlsx, verifies the file exists, accesses the first worksheet, checks for at least one ListObject, renames its DisplayName to a convention‑based identifier (tbl_{WorksheetName}_1), saves the workbook as output.xlsx, and handles missing files or runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (worksheet.ListObjects.Count > 0)
            {
                // Build a meaningful table name following project conventions
                // Example convention: tbl_{WorksheetName}_{TableIndex}
                string newTableName = $"tbl_{worksheet.Name}_1";

                // Rename the first table using the DisplayName property (supported across versions)
                worksheet.ListObjects[0].DisplayName = newTableName;
            }
            else
            {
                Console.WriteLine("No tables (ListObjects) found in the worksheet.");
            }

            // Save the workbook with the updated table name
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
