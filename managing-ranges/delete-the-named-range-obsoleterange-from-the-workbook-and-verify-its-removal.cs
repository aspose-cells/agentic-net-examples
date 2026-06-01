using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class DeleteNamedRangeDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Create a sample range and assign the name "ObsoleteRange"
            AsposeRange range = sheet.Cells.CreateRange("A1:B2");
            range.Name = "ObsoleteRange";

            // Verify the named range exists before removal
            Console.WriteLine("Names count before removal: " + workbook.Worksheets.Names.Count);
            AsposeRange beforeRemoval = workbook.Worksheets.GetRangeByName("ObsoleteRange");
            Console.WriteLine("Range found before removal: " + (beforeRemoval != null));

            // Delete the named range using NameCollection.Remove(string)
            workbook.Worksheets.Names.Remove("ObsoleteRange");

            // Verify the named range has been removed
            Console.WriteLine("Names count after removal: " + workbook.Worksheets.Names.Count);
            AsposeRange afterRemoval = workbook.Worksheets.GetRangeByName("ObsoleteRange");
            Console.WriteLine("Range found after removal: " + (afterRemoval != null));

            // Save the workbook (optional)
            workbook.Save("DeleteNamedRangeDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}