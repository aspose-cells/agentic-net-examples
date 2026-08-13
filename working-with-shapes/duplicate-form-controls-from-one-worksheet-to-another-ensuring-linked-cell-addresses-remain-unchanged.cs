// Title: Copy a worksheet with its form controls while keeping linked cells unchanged – Aspose.Cells for .NET
// Description: Shows how to load an Excel workbook, add a duplicate sheet, and copy the source worksheet using CopyOptions.ReferToDestinationSheet = false so that form‑control references remain on the original sheet. The modified workbook is saved as a new file.
// Keywords: Aspose.Cells | C# worksheet copy | form controls duplication | preserve cell links | CopyOptions | ReferToDestinationSheet | Excel automation | duplicate sheet with shapes | .NET Excel API
// Common Searches: Aspose.Cells copy worksheet without changing form control links | How to keep form control references when duplicating a sheet in C# | CopyOptions ReferToDestinationSheet example | Duplicate Excel sheet with dropdowns preserving source links | C# copy worksheet drawing objects Aspose.Cells
// Developer Intent: Create a copy of an existing worksheet that includes all form controls, ensuring the controls still point to cells on the original sheet.
// Use Cases: Generate per‑user templates that share a single data‑entry range on a master sheet. | Produce multi‑page reports where each page needs the same set of checkboxes and dropdowns, but calculations stay centralized. | Automate workbook setup for different departments while maintaining a single source of truth for linked cell values.
// AI Prompts: Write C# code with Aspose.Cells to duplicate a worksheet and retain the original form‑control cell references. | Explain the effect of CopyOptions.ReferToDestinationSheet on form control links during a sheet copy operation. | Provide a verification step to confirm that copied form controls still reference cells on the source worksheet.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to load an Excel workbook, add a duplicate sheet, and copy the source worksheet using CopyOptions.ReferToDestinationSheet = false so that form‑control references remain on the original sheet. The modified workbook is saved as a new file.
class DuplicateFormControls
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains the source worksheet with form controls
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the source worksheet (change the name as needed)
            Worksheet sourceSheet = workbook.Worksheets["Source"]; // assume a sheet named "Source"
            if (sourceSheet == null)
            {
                Console.WriteLine("Source worksheet 'Source' not found.");
                return;
            }

            // Add a new worksheet that will receive the duplicated controls
            Worksheet destSheet = workbook.Worksheets.Add("Source_Copy");

            // Configure copy options:
            // - ReferToDestinationSheet = false ensures that any linked cell references
            //   (including those used by form controls) continue to point to the original sheet.
            CopyOptions copyOptions = new CopyOptions
            {
                ReferToDestinationSheet = false
            };

            // Copy the source worksheet's contents, formats, and drawing objects (form controls)
            // to the destination worksheet using the specified options.
            destSheet.Copy(sourceSheet, copyOptions);

            // Save the workbook with the duplicated form controls.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
