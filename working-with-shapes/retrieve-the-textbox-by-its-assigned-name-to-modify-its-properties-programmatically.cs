// Title: C# – Retrieve and Modify a Named TextBox in Aspose.Cells
// Description: Shows how to add a TextBox to a worksheet, assign a Name, fetch it through the name indexer, and update its text, font family, size, and bold attribute before saving the workbook.
// Keywords: Aspose.Cells TextBox name | C# retrieve textbox by name | Aspose.Cells modify textbox properties | Aspose.Cells shape name indexer | Aspose.Cells set font size | Aspose.Cells workbook save | Aspose.Cells .NET example
// Common Searches: Aspose.Cells get textbox by name C# | How to change font of a specific TextBox in Aspose.Cells | Retrieve named shape in Excel using Aspose.Cells | Update TextBox text programmatically Aspose.Cells | Access TextBox collection by name Aspose.Cells
// Developer Intent: Find a TextBox using its assigned Name property and programmatically change its content and formatting.
// Use Cases: Replace placeholder text in a template workbook automatically. | Apply corporate font styling to a particular TextBox identified by name. | Adjust the appearance of a chart title stored as a TextBox before distribution. | Batch‑process multiple workbooks to update named TextBoxes with dynamic data.
// AI Prompts: Write C# code with Aspose.Cells to locate a TextBox named 'HeaderBox' and set its background color to LightGray. | Provide an example that retrieves a TextBox by name and adds a hyperlink to its text using Aspose.Cells for .NET. | Create a loop that scans all TextBoxes in a worksheet, selects the one named 'SummaryBox', and changes its horizontal alignment to Center.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to add a TextBox to a worksheet, assign a Name, fetch it through the name indexer, and update its text, font family, size, and bold attribute before saving the workbook.
class RetrieveTextboxByName
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox to the worksheet
        int textBoxIndex = worksheet.TextBoxes.Add(5, 5, 200, 50);
        TextBox textBox = worksheet.TextBoxes[textBoxIndex];

        // Assign a name to the textbox and set initial text
        textBox.Name = "MyTextBox";
        textBox.Text = "Original Text";

        // Retrieve the textbox by its assigned name
        TextBox retrievedTextBox = worksheet.TextBoxes["MyTextBox"];
        if (retrievedTextBox != null)
        {
            // Modify properties of the retrieved textbox
            retrievedTextBox.Text = "Updated Text";
            retrievedTextBox.Font.Name = "Calibri";
            retrievedTextBox.Font.Size = 14;
            retrievedTextBox.Font.IsBold = true;
        }

        // Save the workbook
        workbook.Save("RetrieveTextboxByName.xlsx");
    }
}
