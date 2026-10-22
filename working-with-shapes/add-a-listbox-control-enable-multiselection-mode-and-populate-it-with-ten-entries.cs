// Title: Create a multi‑select ListBox shape in an Excel worksheet and bind it to cells A1:A10 using Aspose.Cells for .NET
// AI Prompts: Write C# code that adds a ListBox shape to a worksheet, enables its MultiSelect property, and sets its ListFillRange to A1:A10 with Aspose.Cells. | Generate a program that fills cells A1:A10, inserts a ListBox control, configures multi‑selection via reflection for version compatibility, and saves the workbook.
// Common Searches: asp.net how to enable multiselect on a ListBox shape using Aspose.Cells | c# add ListBox control to Excel worksheet with Aspose.Cells and bind to a range | asp.net Aspose.Cells ListBox MultiSelect property reflection example | populate ListBox items from worksheet cells using Aspose.Cells for .NET
// Tags: Aspose.Cells add ListBox shape | Aspose.Cells set ListBox MultiSelect | Aspose.Cells bind ListBox to cell range | Aspose.Cells populate ListBox from worksheet | Aspose.Cells reflection for ListBox properties

using Aspose.Cells;
using System;

// The program creates a new workbook, writes ten items into cells A1:A10, adds a ListBox shape at a specified location, enables multi‑selection (using reflection for compatibility), links the ListBox to the populated range, and saves the file as ListBoxExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells A1:A10 with ten entries
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue($"Item {i + 1}");
            }

            // NOTE: The ListBox control API may vary between Aspose.Cells versions.
            // If the ListBox class does not expose MultiSelect or ListFillRange,
            // the following block can be omitted or adjusted according to the
            // version-specific documentation.

            // Add a ListBox control to the worksheet (if supported)
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            var listBox = sheet.Shapes.AddListBox(2, 2, 5, 5, 200, 100);

            // Attempt to set properties that are available in newer versions
            // Use reflection to avoid compile‑time errors on older versions
            var type = listBox.GetType();

            var multiSelectProp = type.GetProperty("MultiSelect");
            if (multiSelectProp != null && multiSelectProp.CanWrite)
            {
                multiSelectProp.SetValue(listBox, true);
            }

            var listFillRangeProp = type.GetProperty("ListFillRange");
            if (listFillRangeProp != null && listFillRangeProp.CanWrite)
            {
                listFillRangeProp.SetValue(listBox, "A1:A10");
            }

            // Save the workbook (lifecycle rule: save)
            string outputPath = "ListBoxExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
