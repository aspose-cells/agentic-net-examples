using System;
using Aspose.Cells;

namespace AsposeCellsDynamicNamedRangeDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and give it a name
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Populate sample data in column A (dynamic range will depend on this data)
                for (int i = 0; i < 10; i++)
                {
                    sheet.Cells[i, 0].PutValue($"Item {i + 1}");
                }

                // Add a dynamic named range "SalesData" using OFFSET and COUNTA
                int nameIndex = workbook.Worksheets.Names.Add("SalesData");
                Name salesName = workbook.Worksheets.Names[nameIndex];
                // The formula creates a range that starts at A1 and expands down to the last non‑empty cell in column A
                salesName.RefersTo = "=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)";

                // Retrieve the range object represented by the named range
                Aspose.Cells.Range salesRange = salesName.GetRange();

                // Get the address of the range and log it
                string address = salesRange.Address;
                Console.WriteLine($"Dynamic named range \"SalesData\" refers to address: {address}");

                // Save the workbook (optional, just to demonstrate lifecycle usage)
                workbook.Save("DynamicNamedRangeDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}