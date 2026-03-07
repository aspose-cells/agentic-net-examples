using System;
using Aspose.Cells;

class ImportJsonSmartMarkers
{
    static void Main()
    {
        // Create a workbook with smart markers
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Employees";

        // Header row
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["C1"].PutValue("Department");

        // Smart marker row
        sheet.Cells["A2"].PutValue("&=$Employees.Name");
        sheet.Cells["B2"].PutValue("&=$Employees.Age");
        sheet.Cells["C2"].PutValue("&=$Employees.Department");

        // JSON data source
        string jsonData = @"{
            ""Employees"": [
                { ""Name"": ""John Doe"", ""Age"": 30, ""Department"": ""Sales"" },
                { ""Name"": ""Jane Smith"", ""Age"": 28, ""Department"": ""Marketing"" }
            ]
        }";

        // Bind data source and process smart markers
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;
        designer.SetJsonDataSource("Employees", jsonData);
        designer.Process();

        // Save the populated workbook
        workbook.Save("ResultFromJsonSmartMarkers.xlsx");
    }
}