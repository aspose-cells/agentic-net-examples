// Title: Aspose.Cells .NET – Retrieve a TextBox by Its Name and Change Formatting
// Description: This example shows how to add a TextBox to a worksheet, assign a custom name, locate the shape using the name indexer (worksheet.TextBoxes["MyTextBox"]), and then update its text, font family, size, bold style, and optional fill color before saving the workbook.
// Keywords: Aspose.Cells retrieve textbox | textbox name indexer .NET | modify textbox font Aspose.Cells | change textbox fill color C# | Aspose.Cells shape by name | update textbox text programmatically | Aspose.Cells TextBox properties | C# Aspose.Cells example | Excel shape manipulation Aspose | Aspose.Cells workbook editing
// Common Searches: How to get a TextBox by name in Aspose.Cells C# | Change font of a specific TextBox in an Excel file using Aspose.Cells | Set background color of a named TextBox with Aspose.Cells .NET | Retrieve and edit a shape by its Name property in Aspose.Cells | Aspose.Cells example for updating TextBox content
// Developer Intent: Locate a TextBox on a worksheet using its assigned Name and programmatically adjust its content and visual attributes.
// Use Cases: Dynamically replace placeholder text in a report label after data is generated. | Apply a corporate font style to a designated TextBox across multiple sheets. | Highlight a particular annotation by changing its fill color once the TextBox is identified by name.
// AI Prompts: Generate C# code with Aspose.Cells that finds a TextBox called "HeaderBox" and sets its background to LightBlue. | Write a loop in C# using Aspose.Cells to iterate all TextBoxes whose names start with "Data_" and change their font to Arial, size 12. | Provide an Aspose.Cells .NET snippet that retrieves a TextBox by name and adds a hyperlink to the displayed text.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example shows how to add a TextBox to a worksheet, assign a custom name, locate the shape using the name indexer (worksheet.TextBoxes["MyTextBox"]), and then update its text, font family, size, bold style, and optional fill color before saving the workbook.
class RetrieveTextboxByName
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a textbox at the specified position and size
            int tbIndex = worksheet.TextBoxes.Add(5, 5, 200, 50);
            TextBox tb = worksheet.TextBoxes[tbIndex];
            tb.Name = "MyTextBox";
            tb.Text = "Original Text";

            // Retrieve the textbox by its assigned name
            TextBox retrieved = worksheet.TextBoxes["MyTextBox"];
            if (retrieved != null)
            {
                // Update textbox properties
                retrieved.Text = "Updated Text";
                retrieved.Font.Name = "Calibri";
                retrieved.Font.Size = 14;
                retrieved.Font.IsBold = true;

                // Set fill color if the API supports it (commented out to avoid compile issues on older versions)
                // retrieved.Fill.ForeColor = Color.Yellow;
            }

            // Save the workbook
            string outputPath = "RetrieveTextboxByName.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
