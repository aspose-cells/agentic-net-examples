// Title: Create master‑detail Excel file with Aspose.Cells smart markers and the DetailSheet parameter in C#
// AI Prompts: Write C# code that adds a master worksheet with smart markers, sets the &DetailSheet parameter, and directs repeated rows to a separate "Detail" sheet using Aspose.Cells. | Show how to bind a List<Person> to the smart marker name "Data", define a named range for the markers, and invoke WorkbookDesigner.Process on that range while preserving unknown markers. | Demonstrate saving the workbook to an XLSX file after processing, ensuring the detail rows appear on the designated worksheet.
// Common Searches: Aspose.Cells C# example using &DetailSheet to write smart marker rows to another worksheet | how to process only a specific smart marker range with WorkbookDesigner in .NET | binding a List<T> to smart markers in Aspose.Cells workbook designer | preserve unrecognized smart markers while processing Aspose.Cells smart markers | generate master‑detail Excel report with smart markers and separate detail sheet in C#
// Tags: Aspose.Cells smart markers DetailSheet parameter | WorkbookDesigner process named range C# | bind List<T> to smart markers Aspose.Cells | output detail rows to separate worksheet Excel | master detail workbook generation Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The sample creates a new workbook, adds a master sheet with header smart markers and the &DetailSheet=Detail marker, defines a named range for the markers, binds a List<Person> as the data source, processes only that range while preserving unknown markers, and saves the result to "DetailSmartMarkers.xlsx" where the repeated detail rows are automatically placed on a separate "Detail" worksheet.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add a master worksheet that will contain the smart markers
            Worksheet masterSheet = workbook.Worksheets[0];
            masterSheet.Name = "Master";

            // Add a separate worksheet that will receive the detail rows
            Worksheet detailSheet = workbook.Worksheets.Add("Detail");

            // ----- Set up smart markers on the master sheet -----
            // Header row
            masterSheet.Cells["A1"].PutValue("Name");
            masterSheet.Cells["B1"].PutValue("Value");

            // DetailSheet parameter tells the designer to place detail rows on the "Detail" sheet
            masterSheet.Cells["A2"].PutValue("&DetailSheet=Detail");

            // Smart markers that will be repeated for each data item
            masterSheet.Cells["A3"].PutValue("&=Data.Name");
            masterSheet.Cells["B3"].PutValue("&=Data.Value");

            // Define the range that contains the smart markers and give it the required name
            Aspose.Cells.Range smartMarkerRange = masterSheet.Cells.CreateRange("A2:B3");
            smartMarkerRange.Name = "_CellsSmartMarkers";

            // ----- Prepare sample data source -----
            List<Person> persons = new List<Person>
            {
                new Person { Name = "Alice", Value = 100 },
                new Person { Name = "Bob",   Value = 200 },
                new Person { Name = "Carol", Value = 300 }
            };

            // ----- Configure WorkbookDesigner and process the smart markers -----
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            // Bind the data source to the name used in the smart markers ("Data")
            designer.SetDataSource("Data", persons);

            // Process only the defined range (the second parameter 'true' preserves unrecognized markers)
            designer.Process(smartMarkerRange, true);

            // Save the result (lifecycle: save)
            workbook.Save("DetailSmartMarkers.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simple data class used as the data source
    public class Person
    {
        public string Name { get; set; } = null!;
        public int Value { get; set; }
    }
}
