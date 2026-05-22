using System;
using System.Text;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using ARange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Register encoding provider (required for .NET Core)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with headers
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("John");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Mary");
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Bob");

            // Add a ListObject (table) covering the range A1:B4, indicating that the range has headers
            int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "EmployeeTable";

            // Retrieve the data range of the table (excludes header row)
            ARange dataRange = table.DataRange;

            // Create a named range that references only the data body of the table
            dataRange.Name = "EmployeeData";

            // Save the workbook
            string outputPath = "TableWithNamedDataRange.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}