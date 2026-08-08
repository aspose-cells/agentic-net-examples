// Title: C# – Import Mixed‑Type ArrayList into Aspose.Cells with Column Number & Date Formatting
// Description: Shows how to create a workbook, import a header and a data row from an ArrayList that holds strings, integers, doubles and DateTime values, then apply column‑specific styles—integer format for Age, two‑decimal format for Salary, and custom yyyy‑MM‑dd format for HireDate—before saving the file.
// Keywords: Aspose.Cells | ImportArrayList | C# | mixed data types | column formatting | integer number format | double format | custom date format | Excel export | ArrayList to worksheet
// Common Searches: Aspose.Cells import ArrayList C# | set column number format Aspose.Cells | apply custom date format to Excel column C# | preserve data types when importing to Aspose.Cells | format Age column as integer Aspose.Cells
// Developer Intent: Load heterogeneous data from an ArrayList into an Excel sheet and enforce appropriate cell formatting for each column using Aspose.Cells.
// Use Cases: Create a report where headers and rows are generated from collections containing mixed types. | Display numeric columns with specific precision—no decimals for integers, two decimals for doubles. | Show dates uniformly with a custom yyyy‑MM‑dd pattern. | Build templates that require exact column styling after bulk data import.
// AI Prompts: Write C# code that uses Aspose.Cells to import an ArrayList with string, int, double, and DateTime values and set column styles for integer, double (two decimals), and custom date format. | Explain how ImportArrayList determines cell types and how column styles can override the default formatting in Aspose.Cells. | Extend the example to import multiple rows, auto‑fit all columns, and protect the worksheet before saving.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsArrayListImportDemo
{
    // Shows how to create a workbook, import a header and a data row from an ArrayList that holds strings, integers, doubles and DateTime values, then apply column‑specific styles—integer format for Age, two‑decimal format for Salary, and custom yyyy‑MM‑dd format for HireDate—before saving the file.
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
                // 1. Prepare an ArrayList with mixed data types (header row)
                // -----------------------------------------------------------------
                ArrayList header = new ArrayList
                {
                    "Name",      // string
                    "Age",       // integer
                    "Salary",    // double
                    "HireDate"   // DateTime
                };

                // Import the header horizontally starting at cell A1 (row 0, column 0)
                cells.ImportArrayList(header, 0, 0, false);

                // -----------------------------------------------------------------
                // 2. Prepare a data row with matching mixed types
                // -----------------------------------------------------------------
                ArrayList row = new ArrayList
                {
                    "John Doe",                     // string
                    30,                             // integer
                    55000.75,                       // double
                    new DateTime(2022, 5, 1)        // DateTime
                };

                // Import the data row horizontally starting at cell A2 (row 1, column 0)
                cells.ImportArrayList(row, 1, 0, false);

                // -----------------------------------------------------------------
                // 3. Define column-specific formatting to ensure correct cell types
                // -----------------------------------------------------------------
                // Column 0 (Name) – default string formatting (no special style needed)

                // Column 1 (Age) – integer number format
                Style intStyle = workbook.CreateStyle();
                intStyle.Number = 0; // No decimal places
                cells.Columns[1].SetStyle(intStyle);

                // Column 2 (Salary) – double number format with two decimal places
                Style doubleStyle = workbook.CreateStyle();
                doubleStyle.Number = 2; // Two decimal places
                cells.Columns[2].SetStyle(doubleStyle);

                // Column 3 (HireDate) – custom date format
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Custom = "yyyy-mm-dd";
                cells.Columns[3].SetStyle(dateStyle);

                // -----------------------------------------------------------------
                // 4. Save the workbook
                // -----------------------------------------------------------------
                workbook.Save("ArrayListImportWithFormatting.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
