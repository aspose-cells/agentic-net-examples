// Title: Show a custom message for an empty collection using the If parameter in Aspose.Cells smart markers (C#)
// Common Searches: how to use the If parameter in Aspose.Cells smart markers to show a message for an empty list in C# | Aspose.Cells smart marker conditional output when the data source collection has no rows | display custom text for empty dataset with WorkbookDesigner smart markers | C# example of handling empty smart marker collection using If clause in Aspose.Cells
// Tags: smart marker if parameter | empty collection handling Aspose.Cells | conditional smart marker output C# | WorkbookDesigner empty dataset message | Aspose.Cells custom empty collection text

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads a template workbook, creates an empty Items list, converts it into a DataSet with a matching DataTable, processes the smart markers via WorkbookDesigner, and saves the resulting workbook. It demonstrates how the If parameter can be used to display a custom message when the collection is empty.
class SmartMarkerIfExample
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Ensure the template file exists before loading.
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file \"{templatePath}\" not found.");
                return;
            }

            // Load the workbook that contains the smart marker template.
            Workbook workbook = new Workbook(templatePath);

            // Prepare the data source with an empty collection.
            var data = new
            {
                Items = new List<string>()
            };

            // Convert the anonymous object to a DataSet for WorkbookDesigner.
            DataSet dataSet = new DataSet();

            // Create a DataTable that matches the smart marker collection name.
            DataTable itemsTable = new DataTable("Items");
            itemsTable.Columns.Add("Item", typeof(string));

            // Populate rows (none in this example because the list is empty).
            foreach (var item in data.Items)
            {
                itemsTable.Rows.Add(item);
            }

            dataSet.Tables.Add(itemsTable);

            // Process the smart markers using WorkbookDesigner.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);
            designer.Process();

            // Save the result workbook.
            workbook.Save(resultPath);
            Console.WriteLine($"Result saved to \"{resultPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
