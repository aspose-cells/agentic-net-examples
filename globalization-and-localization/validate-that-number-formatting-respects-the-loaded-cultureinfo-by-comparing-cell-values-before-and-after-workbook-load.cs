using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsCultureValidation
{
    class Program
    {
        static void Main()
        {
            // Define the culture to be used for formatting (German - uses comma as decimal separator)
            CultureInfo germanCulture = new CultureInfo("de-DE");

            // -------------------- Create workbook --------------------
            Workbook workbook = new Workbook();

            // Apply the culture to the workbook settings
            workbook.Settings.CultureInfo = germanCulture;

            // Access the first worksheet and target cell
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Put a numeric value into the cell
            double numericValue = 1234.56;
            cell.PutValue(numericValue);

            // Create a style with a built‑in decimal format that includes grouping
            Style style = workbook.CreateStyle();
            style.Number = 4; // "#,##0.00"
            cell.SetStyle(style);

            // Save the workbook to a temporary file
            string filePath = "CultureValidation.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // -------------------- Load workbook with culture --------------------
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.CultureInfo = germanCulture; // Ensure the same culture is used during load

            Workbook loadedWorkbook = new Workbook(filePath, loadOptions);
            Cell loadedCell = loadedWorkbook.Worksheets[0].Cells["A1"];

            // Retrieve the formatted string value after loading
            string loadedStringValue = loadedCell.StringValue;

            // Expected formatted string according to German culture
            string expectedStringValue = numericValue.ToString("N2", germanCulture); // "1.234,56"

            // -------------------- Validation --------------------
            Console.WriteLine($"Loaded string value : \"{loadedStringValue}\"");
            Console.WriteLine($"Expected string value: \"{expectedStringValue}\"");

            if (loadedStringValue == expectedStringValue)
            {
                Console.WriteLine("Success: Number formatting respects the loaded CultureInfo.");
            }
            else
            {
                Console.WriteLine("Failure: Number formatting does NOT respect the loaded CultureInfo.");
            }
        }
    }
}