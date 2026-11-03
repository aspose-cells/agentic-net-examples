// Title: How to add list‑type data validation to column S in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add list validation to an entire column | Create dropdown list validation for column S in Excel with Aspose.Cells .NET | Restrict values in Excel column S to a predefined array using Aspose.Cells | Apply list‑type data validation to column index 18 in a C# workbook | Set up data validation list for a column in an existing Excel file with Aspose.Cells
// Tags: Aspose.Cells list validation C# | Excel column data validation Aspose | C# dropdown list validation Aspose.Cells | Validate column S values Aspose.Cells | CellArea range validation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads or creates an Excel workbook, defines an allowed string array (Apple, Banana, Cherry), creates a CellArea that spans the whole column S (zero‑based index 18), adds a list‑type validation using those values, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a workbook with a default sheet
            }

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Allowed values for column S
            string[] allowedValues = new string[] { "Apple", "Banana", "Cherry" };

            // Define the area covering the entire column S (zero‑based column index 18)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 1048575,
                StartColumn = 18,
                EndColumn = 18
            };

            // Add list validation to the defined area
            int validationIndex = sheet.Validations.Add(area);
            Validation validation = sheet.Validations[validationIndex];
            validation.Type = ValidationType.List;
            validation.Operator = OperatorType.Equal;
            validation.Formula1 = "\"" + string.Join(",", allowedValues) + "\"";

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
