using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsWorksheetScopedNamedRange
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a worksheet and give it a specific name
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "DataSheet";

            // Populate some sample data in the worksheet (A1:B3)
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Quantity");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(50);
            dataSheet.Cells["A3"].PutValue("Orange");
            dataSheet.Cells["B3"].PutValue(30);

            // Define a named range with worksheet scope
            int nameIndex = workbook.Worksheets.Names.Add("DataSheet!MyRange");
            Name scopedName = workbook.Worksheets.Names[nameIndex];

            // Set the reference of the named range (absolute reference)
            scopedName.RefersTo = "=DataSheet!$A$1:$B$3";

            // Set the SheetIndex to the worksheet's zero‑based index (clarifies scope)
            scopedName.SheetIndex = dataSheet.Index;

            // Retrieve the range using the Name object
            AsposeRange retrievedRange = scopedName.GetRange();

            // Demonstrate that the range was retrieved correctly
            Console.WriteLine("Retrieved range address: " + retrievedRange.Address);
            Console.WriteLine("Associated worksheet: " + retrievedRange.Worksheet.Name);

            // Optionally, modify the range (e.g., make header text uppercase)
            foreach (Cell cell in retrievedRange)
            {
                if (cell.Row == 0) // first row (header)
                {
                    cell.PutValue(cell.StringValue.ToUpper());
                }
            }

            // Save the workbook in XLSX format
            workbook.Save("WorksheetScopedNamedRange.xlsx");
        }
    }
}