// Title: Extract raw scientific notation strings from Excel cells using GetStringValue with CellValueFormatStrategy.None and convert to double in Aspose.Cells for .NET
// AI Prompts: Call GetStringValue(CellValueFormatStrategy.None) on a cell to obtain the exact text entered, preserving scientific‑notation formatting. | Parse the retrieved string to a double using double.Parse with NumberStyles.Float and CultureInfo.InvariantCulture. | Insert the parsed double into another worksheet cell and save the workbook to persist the numeric values.
// Common Searches: Aspose.Cells C# get raw cell text without formatting for scientific notation | Parse Excel scientific notation string to double using GetStringValue | CellValueFormatStrategy.None example for unformatted values in .NET | Convert Excel scientific notation to numeric value with Aspose.Cells | Read and write double values after extracting raw strings in Aspose.Cells workbook
// Tags: GetStringValue raw text Aspose.Cells | CellValueFormatStrategy.None unformatted cell value | parse scientific notation double C# | convert Excel string to numeric Aspose | write double to worksheet Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates inserting scientific‑notation strings into cells, retrieving the exact text with GetStringValue(CellValueFormatStrategy.None), parsing the strings to double values using invariant culture, writing the numeric results back to other cells, and saving the workbook.
    class GetScientificNotationStringDemo
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Insert scientific notation values as strings
            cells["A1"].PutValue("1.23E+04");   // 12300
            cells["A2"].PutValue("-5.67e-03"); // -0.00567
            cells["A3"].PutValue("9.0E0");     // 9

            // Retrieve the raw string without any formatting using CellValueFormatStrategy.None
            string rawA1 = cells["A1"].GetStringValue(CellValueFormatStrategy.None);
            string rawA2 = cells["A2"].GetStringValue(CellValueFormatStrategy.None);
            string rawA3 = cells["A3"].GetStringValue(CellValueFormatStrategy.None);

            // Convert the scientific notation strings to numeric values
            double valA1 = double.Parse(rawA1, NumberStyles.Float, CultureInfo.InvariantCulture);
            double valA2 = double.Parse(rawA2, NumberStyles.Float, CultureInfo.InvariantCulture);
            double valA3 = double.Parse(rawA3, NumberStyles.Float, CultureInfo.InvariantCulture);

            // Display the results
            Console.WriteLine($"Cell A1 raw string: {rawA1} => numeric: {valA1}");
            Console.WriteLine($"Cell A2 raw string: {rawA2} => numeric: {valA2}");
            Console.WriteLine($"Cell A3 raw string: {rawA3} => numeric: {valA3}");

            // Optionally, write the numeric results back to other cells
            cells["B1"].PutValue(valA1);
            cells["B2"].PutValue(valA2);
            cells["B3"].PutValue(valA3);

            // Save the workbook
            workbook.Save("ScientificNotationDemo.xlsx");
        }
    }
}
