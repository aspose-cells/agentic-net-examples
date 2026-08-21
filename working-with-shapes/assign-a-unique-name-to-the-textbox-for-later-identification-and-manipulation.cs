// Title: Aspose.Cells C# – Assign a Unique Name to a TextBox and Retrieve It by Name
// Description: Shows how to create a workbook, add a TextBox, set its Name property to a unique identifier, save the file, reload it, locate the TextBox using the name indexer, modify its text, and save the updated workbook.
// Keywords: Aspose.Cells TextBox Name | C# named TextBox Aspose | retrieve TextBox by name Aspose.Cells | Aspose.Cells shape naming | Excel TextBox manipulation C# | Aspose.Cells workbook shapes | set TextBox identifier | update TextBox content Aspose | Aspose.Cells .NET TextBox example
// Common Searches: how to name a TextBox in Aspose.Cells C# | retrieve TextBox by name Aspose.Cells .NET | Aspose.Cells change TextBox text after loading workbook | C# Aspose.Cells add and edit TextBox | access Excel shape by name using Aspose.Cells
// Developer Intent: Create a uniquely named TextBox and later modify it via its Name property.
// Use Cases: Form‑like data entry where a specific TextBox must be updated without scanning all shapes. | Version‑controlled annotations that need consistent identification across workbook revisions. | Automation scripts that target a particular TextBox among many shapes for dynamic content updates. | Generating reports where placeholder TextBoxes are replaced with calculated values. | Building Excel‑based UI components that require stable references.
// AI Prompts: Write C# code with Aspose.Cells to add a TextBox, assign a custom Name, save the workbook, reload it, and change the TextBox content using the Name indexer. | Explain the purpose of the TextBox.Name property in Aspose.Cells and best practices for retrieving shapes by name after loading a workbook. | Show how to implement error handling when accessing a TextBox by name that may not exist, including null checks and exception management. | Provide a step‑by‑step guide to rename an existing TextBox in an Excel file using Aspose.Cells C#. | Create a reusable method that accepts a workbook path, TextBox name, and new text, then updates the specified TextBox.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add a TextBox, set its Name property to a unique identifier, save the file, reload it, locate the TextBox using the name indexer, modify its text, and save the updated workbook.
class Program
{
    static void Main()
    {
        // ---------- Create a workbook and add a named TextBox ----------
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Add a TextBox to the worksheet and obtain its index
        int tbIndex = ws.TextBoxes.Add(2, 2, 150, 50);
        TextBox tb = ws.TextBoxes[tbIndex];

        // Assign a unique name for later identification
        tb.Name = "UniqueTextBox1";

        // Set initial text (optional)
        tb.Text = "Initial content";

        // Save the workbook
        wb.Save("NamedTextbox.xlsx");

        // ---------- Load the workbook and manipulate the named TextBox ----------
        Workbook loadedWb = new Workbook("NamedTextbox.xlsx");
        Worksheet loadedWs = loadedWb.Worksheets[0];

        // Retrieve the TextBox by its unique name
        TextBox namedTb = loadedWs.TextBoxes["UniqueTextBox1"];
        if (namedTb != null)
        {
            // Update the text of the TextBox
            namedTb.Text = "Updated content";
        }

        // Save the updated workbook
        loadedWb.Save("NamedTextbox_Updated.xlsx");
    }
}
