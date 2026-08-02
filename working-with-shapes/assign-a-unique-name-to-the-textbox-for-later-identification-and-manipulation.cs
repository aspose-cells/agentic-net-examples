// Title: Assign and Retrieve a Named TextBox Shape with Aspose.Cells for .NET
// Description: Demonstrates how to add a TextBox to a worksheet, set its Name property to a custom identifier, retrieve the same TextBox by name, display its properties, and save the workbook as TextboxNamed.xlsx using C# and Aspose.Cells.
// Keywords: Aspose.Cells TextBox Name | C# set textbox identifier | retrieve textbox by name Aspose.Cells | named shape Aspose.Cells .NET | worksheet textbox manipulation | Aspose.Cells shape naming example
// Common Searches: how to name a textbox in Aspose.Cells | retrieve a textbox by custom name C# | Aspose.Cells set and get TextBox Name property | C# Aspose.Cells example for named shapes | access worksheet textbox using its Name
// Developer Intent: Create a TextBox shape, assign a unique Name for later reference, and fetch the same shape by that Name in a .NET workbook.
// Use Cases: Programmatically locate a specific textbox to update its text, font, or position. | Use named textboxes as placeholders for dynamic data insertion or validation. | Maintain a collection of identifiable shapes when generating reports or templates.
// AI Prompts: Show C# code that assigns a unique Name to a TextBox shape and later retrieves it with Aspose.Cells. | Generate a loop that scans all TextBoxes in a worksheet and returns the one matching a given Name. | Explain how to modify the text and formatting of a named TextBox after it has been retrieved.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a TextBox to a worksheet, set its Name property to a custom identifier, retrieve the same TextBox by name, display its properties, and save the workbook as TextboxNamed.xlsx using C# and Aspose.Cells.
class AssignTextboxName
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox to the worksheet
        int textboxIndex = worksheet.TextBoxes.Add(5, 5, 150, 50);
        TextBox textbox = worksheet.TextBoxes[textboxIndex];

        // Assign a unique name for later identification
        textbox.Name = "MyUniqueTextBox";

        // Example of later retrieval using the assigned name
        TextBox retrievedTextbox = worksheet.TextBoxes["MyUniqueTextBox"];
        if (retrievedTextbox != null)
        {
            Console.WriteLine("Found textbox with name: " + retrievedTextbox.Name);
            Console.WriteLine("Textbox text: " + retrievedTextbox.Text);
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("TextboxNamed.xlsx");
    }
}
