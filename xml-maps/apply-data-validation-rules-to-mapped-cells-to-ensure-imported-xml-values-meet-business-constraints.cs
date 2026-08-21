// Title: C# Sample: Apply Whole‑Number Data Validation to XML‑Mapped Cells Using Aspose.Cells
// Description: Loads a template workbook with CheckDataValid enabled, imports XML into Sheet1, defines a validation range (A2:A100), adds a whole‑number rule (1‑100) with custom messages, and saves the result. Demonstrates XML mapping combined with data validation in Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# data validation | ImportXml Aspose.Cells | CheckDataValid load option | XML mapping validation .NET | whole number validation Aspose | Excel data validation programmatically | Aspose.Cells example GitHub | C# workbook import XML
// Common Searches: Aspose.Cells add numeric range validation after ImportXml | C# load workbook with data validation enabled Aspose | How to enforce whole number rule on XML‑mapped cells | Aspose.Cells example for XML maps and validation | Validate imported XML values in Excel using Aspose
// Developer Intent: Add a numeric range validation rule to cells populated from an XML map and ensure the rule is enforced when the workbook is loaded and saved.
// Use Cases: Guarantee that ID fields imported from XML stay within a permitted range (1‑100). | Prevent users from entering invalid numbers in a pre‑filled template column. | Combine XML mapping with data‑validation to maintain data integrity in automated reporting. | Enable CheckDataValid to let Aspose.Cells raise errors for out‑of‑range values during processing.
// AI Prompts: Write C# code with Aspose.Cells that imports an XML file into column B (rows 5‑200) and applies a whole‑number validation between 10 and 500, including custom input and error messages. | Explain the interaction between the CheckDataValid load option and data validation in Aspose.Cells, and show how to capture validation failures programmatically.

using System;
using Aspose.Cells;
using System.IO;

// Loads a template workbook with CheckDataValid enabled, imports XML into Sheet1, defines a validation range (A2:A100), adds a whole‑number rule (1‑100) with custom messages, and saves the result. Demonstrates XML mapping combined with data validation in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Paths for the template workbook, XML source and output workbook
        string templatePath = "Template.xlsx";
        string xmlPath = "Data.xml";
        string outputPath = "Result.xlsx";

        // ------------------------------------------------------------
        // Load the template workbook with data‑validation checking enabled
        // ------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CheckDataValid = true;               // enforce validation while loading
        Workbook workbook = new Workbook(templatePath, loadOptions);

        // ------------------------------------------------------------
        // Import XML data into the first worksheet starting at cell A1
        // ------------------------------------------------------------
        workbook.ImportXml(xmlPath, "Sheet1", 0, 0);

        // ------------------------------------------------------------
        // Add a data‑validation rule to column A (rows 2‑100)
        // ------------------------------------------------------------
        Worksheet sheet = workbook.Worksheets[0];

        // Define the cell area that the validation will cover
        CellArea validationArea = new CellArea
        {
            StartRow = 1,    // Row 2 (0‑based index)
            StartColumn = 0, // Column A
            EndRow = 99,     // Row 100
            EndColumn = 0    // Column A
        };

        // Add the validation to the worksheet and obtain the Validation object
        int validationIndex = sheet.Validations.Add(validationArea);
        Validation validation = sheet.Validations[validationIndex];

        // Configure the validation: whole numbers between 1 and 100
        validation.Type = ValidationType.WholeNumber;
        validation.Operator = OperatorType.Between;
        validation.Formula1 = "1";
        validation.Formula2 = "100";

        // Optional user‑friendly messages
        validation.InputMessage = "Enter a whole number between 1 and 100.";
        validation.ErrorMessage = "Value must be between 1 and 100.";
        validation.ShowInput = true;
        validation.ShowError = true;
        validation.IgnoreBlank = true;
        validation.InCellDropDown = false;

        // ------------------------------------------------------------
        // Save the workbook with the applied validation rules
        // ------------------------------------------------------------
        workbook.Save(outputPath);
    }
}
