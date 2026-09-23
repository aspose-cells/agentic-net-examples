// Title: Automatically expand an Aspose.Cells named range after inserting rows in a .NET worksheet
// AI Prompts: Generate C# code that inserts a specified number of rows at a given position and then updates an existing named range using an OFFSET‑COUNTA formula with Aspose.Cells. | Show how to programmatically redefine a workbook's named range to stay dynamic after row insertion, including formula construction and saving the file.
// Common Searches: Aspose.Cells C# insert rows and keep named range dynamic | How to refresh Excel named range after adding rows using Aspose.Cells .NET | C# Aspose.Cells OFFSET COUNTA formula for expanding named range automatically
// Tags: Aspose.Cells insert rows dynamic named range | OFFSET COUNTA formula Aspose.Cells | C# update Excel named range after row insertion | Aspose.Cells workbook named range refresh

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads or creates a workbook, inserts three rows at a specified index, rebuilds the "MyRange" named range with an OFFSET‑COUNTA formula that expands based on column A, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Load the workbook; create a new one if the input file does not exist
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    // Create a new workbook with a default worksheet
                    workbook = new Workbook();
                    workbook.Worksheets[0].Name = "Sheet1";
                }

                // Get the first worksheet (adjust index or name as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Insert new rows (e.g., insert 3 rows starting at row index 5 – zero‑based)
                int insertAtRow = 5;          // Row after which new rows will be inserted
                int rowsToInsert = 3;
                sheet.Cells.InsertRows(insertAtRow, rowsToInsert);

                // ----- Update the dynamic named range -----
                // Assume the named range is called "MyRange"
                Name namedRange = workbook.Worksheets.Names["MyRange"];

                // If the named range exists, redefine it using a dynamic formula.
                // Here we use OFFSET together with COUNTA to automatically expand the range
                // based on the number of non‑empty cells in column A.
                if (namedRange != null)
                {
                    // Build the formula: =OFFSET(Sheet1!$A$2,0,0,COUNTA(Sheet1!$A:$A)-1,1)
                    // Adjust the start cell ($A$2) and the column (A) as required for your data.
                    string startCell = "$A$2";
                    string columnLetter = "A";
                    string sheetName = sheet.Name; // e.g., "Sheet1"
                    string dynamicFormula = $"=OFFSET({sheetName}!{startCell},0,0,COUNTA({sheetName}!${columnLetter}:${columnLetter})-1,1)";

                    // Apply the new formula to the named range
                    namedRange.RefersTo = dynamicFormula;
                }

                // Save the modified workbook
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
