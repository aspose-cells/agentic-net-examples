using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DataModels;
using Aspose.Cells.Tables;

namespace RenameTableDemoApp
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with a header row
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("Alice");
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Bob");

            // Add a ListObject (table) covering the data range (including header)
            int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject listObject = worksheet.ListObjects[tableIndex];

            // Rename the table to a meaningful identifier
            listObject.DisplayName = "EmployeeTable";

            // Verify that the DataModel reflects the new name
            if (workbook.DataModel.Tables.Count > 0)
            {
                DataModelTable dataModelTable = workbook.DataModel.Tables[0];
                Console.WriteLine("DataModelTable Name: " + dataModelTable.Name);
            }

            // Save the workbook
            string outputPath = "RenamedTable.xlsx";
            workbook.Save(outputPath);
        }
    }
}