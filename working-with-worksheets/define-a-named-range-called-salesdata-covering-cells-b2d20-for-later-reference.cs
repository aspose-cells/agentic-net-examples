using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (default name is "Sheet1")
            Worksheet sheet = workbook.Worksheets[0];

            // Create a range that covers cells B2 to D20
            AsposeRange salesRange = sheet.Cells.CreateRange("B2", "D20");

            // Assign the name "SalesData" to the created range
            salesRange.Name = "SalesData";

            // (Optional) Verify the named range can be retrieved
            // Range retrieved = workbook.Worksheets.GetRangeByName("SalesData");
            // Console.WriteLine(retrieved.Address);

            // (Optional) Save the workbook if you need to persist the named range
            // string outputPath = "SalesDataWorkbook.xlsx";
            // workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}