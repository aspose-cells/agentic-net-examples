// Title: How to use a master smart marker in an Excel template to repeat rows for each record with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# master smart marker repeat rows for each item in a list | C# generate Excel rows from a collection using WorkbookDesigner and smart markers | populate Excel template with multiple records using Aspose.Cells smart markers | how to bind a list of objects to a master smart marker in Aspose.Cells .NET | repeat template row for each object in C# using Aspose.Cells smart markers
// Tags: master smart marker row duplication | WorkbookDesigner data source collection binding | Aspose.Cells generate rows from list | C# populate Excel template with smart markers | Excel template repeat rows using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example creates or loads an Excel template, inserts master smart markers (&=Master.Name, &=Master.Age) in a row, binds a List<object> as the "Master" data source, processes the markers with WorkbookDesigner to duplicate the row for each record, and saves the populated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string templatePath = "Template.xlsx";

            // Ensure the template file exists; create a simple one if missing
            if (!File.Exists(templatePath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Name = "Sheet1";
                // Header row (optional)
                ws.Cells["A1"].PutValue("Name");
                ws.Cells["B1"].PutValue("Age");
                wb.Save(templatePath);
                Console.WriteLine($"Created placeholder template at '{templatePath}'.");
            }

            // Load the template workbook
            var workbook = new Workbook(templatePath);
            var sheet = workbook.Worksheets[0];

            // Insert smart markers in the template row (starting at A2)
            sheet.Cells["A2"].PutValue("&=Master.Name");
            sheet.Cells["B2"].PutValue("&=Master.Age");

            // Prepare data source: a list of master records
            var masterData = new List<object>
            {
                new { Name = "John", Age = 30 },
                new { Name = "Jane", Age = 25 },
                new { Name = "Bob", Age = 40 }
            };

            // Process smart markers using WorkbookDesigner
            var designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Master", masterData);
            designer.Process();

            // Save the resulting workbook
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
