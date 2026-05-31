using System;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

class VerifyTableConversion
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the table
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("John");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Mary");

            // Add a ListObject (table) covering the data range
            int tableIndex = sheet.ListObjects.Add("A1", "B3", true);
            ListObject listObject = sheet.ListObjects[tableIndex];

            // Verify that the ListObject exists (it is a table at this point)
            Console.WriteLine("Before conversion - ListObject count: " + sheet.ListObjects.Count);

            // Convert the table back to a normal range
            listObject.ConvertToRange();

            // After conversion the ListObject collection should be empty
            bool tableStillExists = sheet.ListObjects.Count > 0;
            Console.WriteLine("After conversion - ListObject exists in collection: " + tableStillExists);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}