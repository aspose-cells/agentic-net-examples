// Title: Aspose.Cells .NET – Build a C# Dictionary of Worksheet Names to Empty‑Flag
// Description: Load a workbook with Aspose.Cells, iterate its worksheets, and use Cells.MaxDataRow / MaxDataColumn (‑1 when no data) to decide if a sheet is empty. The result is stored in a Dictionary<string,bool> and can be printed or saved.
// Keywords: Aspose.Cells empty worksheet detection | C# worksheet MaxDataRow | C# worksheet MaxDataColumn | dictionary worksheet name boolean | check if sheet has data Aspose.Cells | iterate worksheets Aspose.Cells .NET | used range empty sheet C# | Aspose.Cells workbook analysis
// Common Searches: How to tell if a worksheet is empty with Aspose.Cells | Create map of sheet name to empty status in C# | Aspose.Cells MaxDataRow returns -1 for blank sheet | C# code to list empty worksheets in a workbook | Generate dictionary of worksheet emptiness Aspose.Cells
// Developer Intent: Identify every worksheet that contains no data and record its name with a true/false flag.
// Use Cases: Skip blank sheets during bulk data import or export. | Produce a validation report highlighting unintended empty tabs. | Automate cleanup by removing or archiving worksheets that are empty.
// AI Prompts: Write C# code using Aspose.Cells that returns a Dictionary<string,bool> where each key is a worksheet name and the value is true if the sheet has no data, based on MaxDataRow/MaxDataColumn. | Suggest an alternative approach (e.g., using the worksheet's UsedRange or cell count) to detect empty sheets and update the dictionary accordingly.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Load a workbook with Aspose.Cells, iterate its worksheets, and use Cells.MaxDataRow / MaxDataColumn (‑1 when no data) to decide if a sheet is empty. The result is stored in a Dictionary<string,bool> and can be printed or saved.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Dictionary to store worksheet name -> isEmpty mapping
        Dictionary<string, bool> sheetEmptyMap = new Dictionary<string, bool>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // A worksheet is considered empty when it has no data rows or columns.
            // MaxDataRow and MaxDataColumn return -1 when there is no data.
            bool isEmpty = sheet.Cells.MaxDataRow < 0 || sheet.Cells.MaxDataColumn < 0;

            sheetEmptyMap.Add(sheet.Name, isEmpty);
        }

        // Output the results (optional)
        foreach (var kvp in sheetEmptyMap)
        {
            Console.WriteLine($"{kvp.Key}: {(kvp.Value ? "Empty" : "Not Empty")}");
        }

        // Save the workbook if any changes were made (optional)
        workbook.Save("output.xlsx");
    }
}
