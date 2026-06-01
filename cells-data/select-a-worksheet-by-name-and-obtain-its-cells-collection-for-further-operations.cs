using System;
using Aspose.Cells;

namespace Example
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (contains a default worksheet)
            Workbook workbook = new Workbook();

            // Rename the default worksheet
            workbook.Worksheets[0].Name = "DataSheet";

            // Add another worksheet with a specific name
            workbook.Worksheets.Add("Report");

            // Select the worksheet by its name
            Worksheet reportSheet = workbook.Worksheets["Report"];

            // Get the Cells collection of the selected worksheet
            Cells cells = reportSheet.Cells;

            // Example operation: write a value to cell A1
            cells["A1"].PutValue("Hello from Report sheet");

            // Save the workbook
            workbook.Save("SelectWorksheetByName.xlsx", SaveFormat.Xlsx);
        }
    }
}