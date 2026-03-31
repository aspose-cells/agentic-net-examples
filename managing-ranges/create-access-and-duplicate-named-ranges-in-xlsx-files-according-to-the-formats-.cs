using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace NamedRangeDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Populate some sample data in A1:B3
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(15);

            // -------------------------------------------------
            // Create a named range "OriginalRange" that refers to A1:B3
            // -------------------------------------------------
            NameCollection names = workbook.Worksheets.Names;
            int originalIndex = names.Add("OriginalRange");
            Name originalName = names[originalIndex];
            // RefersTo must start with '='
            originalName.RefersTo = $"={sheet.Name}!$A$1:$B$3";

            // -------------------------------------------------
            // Access the range using Name.GetRange()
            // -------------------------------------------------
            AsposeRange originalRange = originalName.GetRange();
            Console.WriteLine($"Original range address: {originalRange.Address}");

            // -------------------------------------------------
            // Duplicate the named range:
            // Create a new name "DuplicateRange" and set its RefersTo
            // to the same reference as the original.
            // -------------------------------------------------
            int duplicateIndex = names.Add("DuplicateRange");
            Name duplicateName = names[duplicateIndex];
            // Use the original name's RefersTo property to copy the reference
            duplicateName.RefersTo = originalName.RefersTo;

            // Verify the duplicated range
            AsposeRange duplicateRange = duplicateName.GetRange();
            Console.WriteLine($"Duplicate range address: {duplicateRange.Address}");

            // -------------------------------------------------
            // Access a named range via WorksheetCollection.GetRangeByName
            // -------------------------------------------------
            AsposeRange fetchedByCollection = workbook.Worksheets.GetRangeByName("OriginalRange");
            if (fetchedByCollection != null)
            {
                Console.WriteLine($"Fetched by collection address: {fetchedByCollection.Address}");
            }

            // -------------------------------------------------
            // Modify the original range (e.g., uppercase header row)
            // -------------------------------------------------
            foreach (Cell cell in originalRange)
            {
                if (cell.Row == 0) // first row (header)
                {
                    cell.PutValue(cell.StringValue.ToUpper());
                }
            }

            // -------------------------------------------------
            // Save the workbook (lifecycle save)
            // -------------------------------------------------
            workbook.Save("NamedRangeDuplicateDemo.xlsx");
        }
    }
}