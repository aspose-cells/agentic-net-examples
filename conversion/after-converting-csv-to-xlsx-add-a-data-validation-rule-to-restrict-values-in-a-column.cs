// Title: Convert a CSV file to XLSX and apply a list‑type data validation dropdown to a column using Aspose.Cells for .NET
// Common Searches: C# how to convert CSV to XLSX and add a drop‑down list validation with Aspose.Cells | Aspose.Cells example for adding list validation to a specific range after loading a CSV | Create data validation list in Excel column after CSV import using Aspose.Cells .NET | Set in‑cell dropdown values for B2:B100 in a workbook generated from CSV with Aspose.Cells | Aspose.Cells C# code to load CSV, define validation range, and save as XLSX
// Tags: CSV to XLSX conversion using Aspose.Cells | add list validation to Excel column with Aspose.Cells | in‑cell dropdown list validation C# | define validation range B2:B100 Aspose.Cells | save workbook as XLSX Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a CSV file into an Aspose.Cells Workbook, creates a validation area (B2:B100), adds a list‑type data validation with the values "Apple,Banana,Cherry", configures dropdown and error settings, and saves the result as an XLSX file.
class CsvToXlsxWithValidation
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.csv";
            const string outputPath = "output.xlsx";
            const string validationRange = "B2:B100";

            // Verify that the input CSV file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the CSV file into a workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Get the first worksheet (the CSV data is placed here)
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare the validation area (e.g., "B2:B100")
            string[] rangeParts = validationRange.Split(':');
            if (rangeParts.Length != 2)
                throw new ArgumentException("Invalid validation range format.");

            CellArea area = CellArea.CreateCellArea(rangeParts[0], rangeParts[1]);

            // Add a new data validation rule to the worksheet using the area
            ValidationCollection validations = sheet.Validations;
            int index = validations.Add(area);
            Validation validation = validations[index];

            // Set the validation type to a list (restrict values to a predefined list)
            validation.Type = ValidationType.List;

            // Provide the list of allowed values (comma‑separated, enclosed in double quotes)
            validation.Formula1 = "\"Apple,Banana,Cherry\"";

            // Optional settings
            validation.IgnoreBlank = true;          // Allow blank cells
            validation.InCellDropDown = true;       // Show drop‑down arrow in cells
            validation.ShowError = true;            // Show error message when invalid data is entered
            validation.ErrorTitle = "Invalid Entry";
            validation.ErrorMessage = "Please select a value from the list.";

            // Save the workbook as an XLSX file
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
