using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsDemo
{
    class DisableAddRowsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the table
                worksheet.Cells["A1"].PutValue("ID");
                worksheet.Cells["B1"].PutValue("Name");
                worksheet.Cells["A2"].PutValue(1);
                worksheet.Cells["B2"].PutValue("Alice");
                worksheet.Cells["A3"].PutValue(2);
                worksheet.Cells["B3"].PutValue("Bob");

                // Add a ListObject (table) covering the data range
                int tableIndex = worksheet.ListObjects.Add("A1", "B3", true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Access protection settings
                Protection protection = worksheet.Protection;

                // Disallow inserting rows while the sheet is protected
                protection.AllowInsertingRow = false;

                // Set a password and protect the worksheet
                protection.Password = "pwd123";
                worksheet.Protect(ProtectionType.All);

                // Define output file path
                string outputPath = "FixedTable.xlsx";

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            DisableAddRowsDemo.Run();
        }
    }
}