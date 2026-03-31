using System;
using Aspose.Cells;

namespace FindOptionsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data and formulas
            cells["A1"].PutValue("Apple");
            cells["A2"].PutValue("Banana");
            cells["A3"].PutValue("Cherry");
            cells["B1"].Formula = "=A1 & \"_Fruit\"";   // Formula containing text
            cells["B2"].Formula = "=SUM(1,2,3)";       // Numeric formula
            cells["B3"].PutValue("[Data_123]");        // Text that matches a regex pattern

            // -------------------------------------------------
            // 1. Find a cell containing the exact value "Banana"
            // -------------------------------------------------
            FindOptions optExact = new FindOptions
            {
                LookInType = LookInType.Values,          // Search in cell values
                LookAtType = LookAtType.EntireContent    // Exact match
            };
            Cell foundExact = cells.Find("Banana", null, optExact);
            Console.WriteLine(foundExact != null
                ? $"Exact value found at {foundExact.Name}"
                : "Exact value not found");

            // -------------------------------------------------
            // 2. Find a cell whose formula contains the text "Fruit"
            // -------------------------------------------------
            FindOptions optFormula = new FindOptions
            {
                LookInType = LookInType.OnlyFormulas,    // Search only in formulas
                LookAtType = LookAtType.Contains          // Partial match
            };
            Cell foundFormula = cells.Find("Fruit", null, optFormula);
            Console.WriteLine(foundFormula != null
                ? $"Formula containing 'Fruit' found at {foundFormula.Name}"
                : "Formula containing 'Fruit' not found");

            // -------------------------------------------------
            // 3. Find cells within a specific range (A1:B2) that contain "Apple"
            // -------------------------------------------------
            FindOptions optRange = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.Contains
            };
            // Define the search area A1:B2
            CellArea area = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 1,
                EndColumn = 1
            };
            optRange.SetRange(area);
            Cell foundInRange = cells.Find("Apple", null, optRange);
            Console.WriteLine(foundInRange != null
                ? $"Value 'Apple' found within range at {foundInRange.Name}"
                : "Value 'Apple' not found in the specified range");

            // -------------------------------------------------
            // 4. Find a cell using a regular expression pattern
            // -------------------------------------------------
            FindOptions optRegex = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.Contains,
                RegexKey = true                         // Enable regex search
            };
            string regexPattern = @"\[[A-Za-z0-9_]+\]"; // Matches strings like [Data_123]
            Cell foundRegex = cells.Find(regexPattern, null, optRegex);
            Console.WriteLine(foundRegex != null
                ? $"Regex match found at {foundRegex.Name} with value '{foundRegex.StringValue}'"
                : "No cell matched the regex pattern");

            // -------------------------------------------------
            // 5. Find a cell by style (e.g., bold font)
            // -------------------------------------------------
            // Create a style with bold font and apply it to a cell
            Style boldStyle = workbook.CreateStyle();
            boldStyle.Font.IsBold = true;
            cells["C1"].PutValue("BoldText");
            cells["C1"].SetStyle(boldStyle);

            FindOptions optStyle = new FindOptions
            {
                Style = boldStyle,                      // Search for cells with this style
                LookInType = LookInType.Values,
                LookAtType = LookAtType.EntireContent
            };
            Cell foundStyle = cells.Find(null, null, optStyle);
            Console.WriteLine(foundStyle != null
                ? $"Cell with matching style found at {foundStyle.Name}"
                : "No cell with the specified style found");

            // Save the workbook (lifecycle rule: save)
            workbook.Save("FindOptionsDemo.xlsx");
        }
    }
}