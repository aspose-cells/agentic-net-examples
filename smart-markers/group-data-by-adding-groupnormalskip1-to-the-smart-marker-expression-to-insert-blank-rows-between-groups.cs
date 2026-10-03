// Title: How to insert blank rows between grouped rows using Aspose.Cells smart markers (group:normal,skip:1) in C#
// AI Prompts: Generate C# code that creates an Excel workbook with Aspose.Cells, defines a smart marker `${Data.Group:normal,skip:1}` to group items and automatically adds an empty row after each group, then binds a List<Item> and processes the markers. | Show how to configure WorkbookDesigner to use a List<T> data source and apply the group:normal,skip:1 option in smart markers to produce grouped sections with separator rows in the output file.
// Common Searches: Aspose.Cells smart marker group normal skip 1 blank row between groups C# | C# insert empty row after each group using WorkbookDesigner smart markers | How to add separator rows when grouping data with Aspose.Cells smart markers | Example of group:normal,skip:1 in Aspose.Cells smart markers for .NET
// Tags: Aspose.Cells smart marker group normal skip 1 | C# WorkbookDesigner group separator rows | Excel export blank rows between groups Aspose.Cells | binding List<T> to smart markers Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, places a smart marker `${Data.Group:normal,skip:1}` to generate grouped rows with a blank row between each group, supplies a List<Item> as the data source, processes the markers via WorkbookDesigner, and saves the file as GroupedData.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Item Name");
            sheet.Cells["C1"].PutValue("Category");

            // Smart marker rows with grouping and a blank row between groups
            sheet.Cells["A2"].PutValue("${Data.Group:normal,skip:1}.Category"); // Group header
            sheet.Cells["B2"].PutValue("${Data.Name}");                           // Item name
            sheet.Cells["C2"].PutValue("${Data.Category}");                      // Category value (optional)

            // Prepare the data source
            List<Item> items = new List<Item>()
            {
                new Item { Name = "Apple",    Category = "Fruit" },
                new Item { Name = "Banana",   Category = "Fruit" },
                new Item { Name = "Carrot",   Category = "Vegetable" },
                new Item { Name = "Broccoli", Category = "Vegetable" },
                new Item { Name = "Chicken",  Category = "Meat" }
            };

            // Process the smart markers using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Data", items); // Correct overload: name + data source
            designer.Process();

            // Save the workbook
            string outputPath = "GroupedData.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simple POCO class representing each row of data
    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
    }
}
