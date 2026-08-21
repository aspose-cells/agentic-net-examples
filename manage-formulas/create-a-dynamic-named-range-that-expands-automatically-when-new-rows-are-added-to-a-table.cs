// Title: Create an auto‑expanding dynamic named range with OFFSET in Aspose.Cells for .NET
// Description: This example shows how to build a workbook, add initial data, define a dynamic named range using an OFFSET‑COUNTA formula, retrieve its address, append new rows, recalculate formulas so the range expands, and finally save the file. The code demonstrates real‑time updating of the named range as data grows.
// Keywords: Aspose.Cells | .NET | C# | dynamic named range | OFFSET formula | COUNTA | auto expand range | add rows to Excel | recalculate formulas | Excel automation | named range update
// Common Searches: Aspose.Cells create dynamic named range | OFFSET formula for expanding range in C# | update named range after inserting rows Aspose.Cells | how to recalculate formulas in Aspose.Cells workbook | C# example of auto‑growing Excel range
// Developer Intent: Define a named range that automatically grows when new rows are added.
// Use Cases: Drive a chart that automatically includes newly added data points. | Apply data validation lists that adjust as the source column expands. | Reference the range in summary calculations (SUM, AVERAGE) that stay current with added rows. | Link the dynamic range to a pivot table so the source updates without manual intervention.
// AI Prompts: Generate C# code using Aspose.Cells to create a dynamic named range for column B with a header row. | Show how to modify the OFFSET formula to ignore blank cells and cap the range at 100 rows. | Provide an example of connecting a dynamic named range to a pivot table in Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace DynamicNamedRangeDemo
{
    // This example shows how to build a workbook, add initial data, define a dynamic named range using an OFFSET‑COUNTA formula, retrieve its address, append new rows, recalculate formulas so the range expands, and finally save the file. The code demonstrates real‑time updating of the named range as data grows.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate initial data in column A (A1 is header)
                cells["A1"].PutValue("Header");
                cells["A2"].PutValue(10);
                cells["A3"].PutValue(20);
                cells["A4"].PutValue(30);

                // Define a dynamic named range that expands with non‑empty rows in column A
                // Formula: =OFFSET(Sheet1!$A$2,0,0,COUNTA(Sheet1!$A:$A)-1,1)
                int nameIdx = workbook.Worksheets.Names.Add("MyDynamicRange");
                Name dynName = workbook.Worksheets.Names[nameIdx];
                dynName.RefersTo = "=OFFSET(Sheet1!$A$2,0,0,COUNTA(Sheet1!$A:$A)-1,1)";

                // Retrieve and display the current range addressed by the dynamic name
                AsposeRange currentRange = dynName.GetRange();
                Console.WriteLine("Initial dynamic range address: " + currentRange.Address);

                // Add new rows of data below the existing data
                cells["A5"].PutValue(40);
                cells["A6"].PutValue(50);
                cells["A7"].PutValue(60);

                // Recalculate formulas so that COUNTA updates
                workbook.CalculateFormula();

                // Retrieve the updated range and display its new address
                AsposeRange updatedRange = dynName.GetRange();
                Console.WriteLine("Updated dynamic range address after adding rows: " + updatedRange.Address);

                // Save the workbook
                workbook.Save("DynamicNamedRangeDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
