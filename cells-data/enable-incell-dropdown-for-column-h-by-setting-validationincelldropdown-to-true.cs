// Title: How to add an in‑cell dropdown list to the entire column H using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create list validation for whole column H in Excel | Enable in‑cell dropdown arrow for column H with Aspose.Cells .NET | Set Validation.InCellDropDown true for Excel column using Aspose.Cells C# | Generate workbook with dropdown list in column H using Aspose.Cells | C# Aspose.Cells add data validation list to column H
// Tags: Aspose.Cells column H dropdown list | C# Excel list validation using Aspose | InCellDropDown property Aspose.Cells | create data validation for entire column C# | Aspose.Cells workbook generation with validation

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, defines the full H column range (H1:H1048576), adds a list‑type validation with predefined options, enables the in‑cell dropdown arrow, ensures the output directory exists, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for column H (all rows)
            // Using the maximum number of rows in Excel (1,048,576)
            CellArea range = CellArea.CreateCellArea("H1", "H1048576");

            // Add a list validation for the defined range
            int validationIndex = sheet.Validations.Add(range);
            Validation validation = sheet.Validations[validationIndex];

            // Set validation type to List
            validation.Type = ValidationType.List;

            // Define allowed values (comma‑separated, enclosed in double quotes)
            validation.Formula1 = "\"Option1,Option2,Option3\"";

            // Enable the in‑cell dropdown arrow
            validation.InCellDropDown = true;

            // Determine output file path
            string outputPath = "Output.xlsx";

            // Ensure the directory exists (handle case when outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
