using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsArrayListImportDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet's cells collection
                Workbook workbook = new Workbook();
                Cells cells = workbook.Worksheets[0].Cells;

                // -----------------------------------------------------------------
                // 1. Prepare header row (string values) and import it horizontally
                // -----------------------------------------------------------------
                ArrayList header = new ArrayList
                {
                    "Name",      // Text
                    "Age",       // Integer
                    "Salary",    // Double
                    "HireDate",  // DateTime
                    "Active"     // Boolean
                };

                // Import header at row 0, column 0 (A1) horizontally (isVertical = false)
                cells.ImportArrayList(header, 0, 0, false);

                // ---------------------------------------------------------------
                // 2. Prepare data rows with mixed types and import them
                // ---------------------------------------------------------------
                // Row 1
                ArrayList row1 = new ArrayList
                {
                    "John Doe",
                    30,
                    55000.75,
                    new DateTime(2022, 5, 1),
                    true
                };
                cells.ImportArrayList(row1, 1, 0, false); // starts at A2

                // Row 2
                ArrayList row2 = new ArrayList
                {
                    "Jane Smith",
                    28,
                    62000.00,
                    new DateTime(2021, 11, 15),
                    false
                };
                cells.ImportArrayList(row2, 2, 0, false); // starts at A3

                // Row 3
                ArrayList row3 = new ArrayList
                {
                    "Bob Johnson",
                    45,
                    72000.5,
                    new DateTime(2020, 2, 20),
                    true
                };
                cells.ImportArrayList(row3, 3, 0, false); // starts at A4

                // ---------------------------------------------------------------
                // 3. Apply column-specific formatting to ensure correct cell types
                // ---------------------------------------------------------------
                // Column B (Age) – integer format (no decimal places)
                Style intStyle = workbook.CreateStyle();
                intStyle.Number = 0; // integer without decimals
                StyleFlag intFlag = new StyleFlag { All = true };
                cells.Columns[1].ApplyStyle(intStyle, intFlag);

                // Column C (Salary) – two decimal places
                Style doubleStyle = workbook.CreateStyle();
                doubleStyle.Number = 2; // number with two decimal places
                StyleFlag doubleFlag = new StyleFlag { All = true };
                cells.Columns[2].ApplyStyle(doubleStyle, doubleFlag);

                // Column D (HireDate) – custom date format
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Custom = "yyyy-mm-dd";
                StyleFlag dateFlag = new StyleFlag { All = true };
                cells.Columns[3].ApplyStyle(dateStyle, dateFlag);

                // Column E (Active) – display as TRUE/FALSE (default text is fine)

                // ---------------------------------------------------------------
                // 4. Save the workbook
                // ---------------------------------------------------------------
                workbook.Save("ArrayListMixedTypes.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}