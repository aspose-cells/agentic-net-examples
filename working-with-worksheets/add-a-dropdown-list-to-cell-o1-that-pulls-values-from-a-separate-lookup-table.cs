// Title: Create an in‑cell drop‑down list in cell O1 using a lookup range in column A with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to add a list‑type validation to cell O1, pulling values from the range A1:A5. | Generate a complete Aspose.Cells example that populates a lookup table in column A, defines the range, and attaches an in‑cell dropdown to a specified cell. | Provide a C# snippet that creates a workbook, sets Validation.Type = List, enables InCellDropDown, and saves the file after adding the dropdown.
// Common Searches: Aspose.Cells C# how to add a data validation list to a single cell | C# create Excel dropdown in cell O1 referencing column A values using Aspose.Cells | Set in‑cell drop‑down list from a lookup range with Aspose.Cells .NET | Aspose.Cells validation formula list example for Excel workbook
// Tags: Aspose.Cells list validation in Excel | C# add in‑cell dropdown with Aspose.Cells | reference lookup range for data validation Aspose.Cells | create workbook with data validation Aspose.Cells C# | save Excel file after adding validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a new workbook, fills cells A1‑A5 with lookup items, defines the range $A$1:$A$5, adds a list‑type validation with an in‑cell drop‑down to cell O1 referencing that range, and saves the workbook as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (or add one if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Populate a lookup table (example values) in column A
                sheet.Cells["A1"].PutValue("Item1");
                sheet.Cells["A2"].PutValue("Item2");
                sheet.Cells["A3"].PutValue("Item3");
                sheet.Cells["A4"].PutValue("Item4");
                sheet.Cells["A5"].PutValue("Item5");

                // Define the range that contains the lookup values
                string lookupRange = "$A$1:$A$5";

                // Add a drop‑down list to cell O1 (column 14, row 0)
                CellArea dropDownArea = new CellArea
                {
                    StartRow = 0,      // Row 1 (0‑based index)
                    EndRow = 0,
                    StartColumn = 14, // Column O (0‑based index)
                    EndColumn = 14
                };

                int validationIndex = sheet.Validations.Add(dropDownArea);
                Validation validation = sheet.Validations[validationIndex];
                validation.Type = ValidationType.List;   // List type validation
                validation.InCellDropDown = true;        // Show the drop‑down arrow
                validation.Formula1 = lookupRange;       // Reference to the lookup table

                // Ensure the output directory exists
                string outputPath = "output.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
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
