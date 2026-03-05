using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook (uses the provided Workbook() constructor)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data with headers
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["A2"].PutValue(10);
        worksheet.Cells["B2"].PutValue(20);
        worksheet.Cells["A3"].PutValue(30);
        worksheet.Cells["B3"].PutValue(40);

        // Add a ListObject (table) using the string‑range overload of ListObjects.Add
        // Parameters: startCell, endCell, hasHeaders
        int listIndex = worksheet.ListObjects.Add("A1", "B3", true);

        // Retrieve the created ListObject for further configuration
        ListObject listObject = worksheet.ListObjects[listIndex];
        listObject.DisplayName = "SampleTable";
        listObject.TableStyleType = TableStyleType.TableStyleMedium2;

        // Save the workbook as an XLSX file (uses the provided Save method)
        workbook.Save("ListObjectDemo.xlsx", SaveFormat.Xlsx);
    }
}