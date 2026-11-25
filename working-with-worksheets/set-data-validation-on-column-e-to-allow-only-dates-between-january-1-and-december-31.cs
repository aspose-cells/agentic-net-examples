// Title: Create date‑range data validation for column E (E1:E1000) with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to restrict column E to dates between Jan 1 and Dec 31 | C# add Excel data validation for a date range using Aspose.Cells | Set date validation on a specific column in an Excel workbook with Aspose.Cells .NET | Apply Between operator for date validation in an Aspose.Cells worksheet | Example of using Validation.Type Date in Aspose.Cells C#
// Tags: Aspose.Cells date validation C# | Excel column E data validation Aspose.Cells | C# set validation between dates Aspose.Cells | Worksheet validation date range .NET | Aspose.Cells ValidationType Date example

using System;
using Aspose.Cells;

// The sample creates a new workbook, defines the cell area covering column E rows 1‑1000, adds a Validation of type Date with the Between operator, sets the allowed dates to 1/1/2023 and 12/31/2023, configures user messages, and saves the file as DateValidation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range for column E (zero‑based column index 4)
            // Here we apply the validation to rows 0 through 999 (E1:E1000)
            CellArea dateRange = CellArea.CreateCellArea(0, 4, 999, 4);

            // Add a new validation to the worksheet for the defined range
            int validationIndex = worksheet.Validations.Add(dateRange);
            Validation validation = worksheet.Validations[validationIndex];

            // Configure the validation to allow only dates between Jan 1 and Dec 31 (any year)
            validation.Type = ValidationType.Date;
            validation.Operator = OperatorType.Between;
            // Use a reference year; Excel will treat the values as dates.
            validation.Formula1 = "1/1/2023";
            validation.Formula2 = "12/31/2023";

            // Optional: user‑friendly messages
            validation.InputTitle = "Enter a Date";
            validation.InputMessage = "Please enter a date between Jan 1 and Dec 31.";
            validation.ErrorTitle = "Invalid Date";
            validation.ErrorMessage = "The date must be between Jan 1 and Dec 31.";
            validation.ShowInput = true;
            validation.ShowError = true;

            // Save the workbook
            string outputPath = "DateValidation.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
