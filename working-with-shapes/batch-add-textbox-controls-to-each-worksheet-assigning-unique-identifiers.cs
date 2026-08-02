// Title: Add TextBox controls with unique names to every worksheet using Aspose.Cells for .NET (C#)
// Description: This C# example creates a new workbook, removes the default sheet, adds three worksheets, and loops through each sheet to insert a TextBox at cell B2 (100 × 200 px). Every TextBox receives a distinct Name built from the worksheet index and shape index, and its Text property displays that name. The workbook is then saved as "BatchTextBoxes.xlsx".
// Keywords: Aspose.Cells | C# | .NET | Excel TextBox | batch add shapes | unique shape name | worksheet TextBox | add TextBox programmatically | Aspose.Cells API example | GitHub sample code
// Common Searches: Aspose.Cells add TextBox to each worksheet | C# batch create TextBox shapes in Excel | assign unique names to Aspose.Cells TextBoxes | how to loop worksheets and add TextBox Aspose.Cells | sample code for adding TextBox controls with Aspose.Cells .NET
// Developer Intent: Insert a TextBox into every worksheet and give each one a unique identifier.
// Use Cases: Generate a template where every sheet contains a labeled instruction box that can be referenced later. | Create placeholder comment boxes across multiple sheets for downstream data‑entry automation. | Automate the addition of identifiable shape objects that can be programmatically updated or removed in subsequent processing.
// AI Prompts: Write a C# snippet using Aspose.Cells to add a TextBox to each worksheet, assign a unique Name based on the sheet index, set its Text, and save the workbook. | Show how to clear default worksheets, add several sheets, batch‑insert TextBoxes with distinct identifiers, and handle exceptions during saving. | Explain how to retrieve and modify a specific TextBox in a saved workbook by using the unique Name assigned during batch creation.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsBatchTextBox
{
    // This C# example creates a new workbook, removes the default sheet, adds three worksheets, and loops through each sheet to insert a TextBox at cell B2 (100 × 200 px). Every TextBox receives a distinct Name built from the worksheet index and shape index, and its Text property displays that name. The workbook is then saved as "BatchTextBoxes.xlsx".
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Remove the default worksheet to avoid duplicate name errors
                if (workbook.Worksheets.Count > 0)
                {
                    workbook.Worksheets.Clear();
                }

                // Add worksheets with unique names
                workbook.Worksheets.Add("Sheet1");
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Iterate through each worksheet and add a TextBox with a unique name
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Add a TextBox at row 1, column 1 with height 100px and width 200px
                    int textBoxIndex = ws.TextBoxes.Add(1, 1, 100, 200);

                    // Retrieve the added TextBox
                    TextBox textBox = ws.TextBoxes[textBoxIndex];

                    // Assign a unique identifier using the worksheet index and TextBox index
                    textBox.Name = $"TextBox_W{ws.Index}_T{textBoxIndex}";

                    // Optional: set some display text
                    textBox.Text = $"This is {textBox.Name}";
                }

                // Save the workbook (lifecycle: save)
                workbook.Save("BatchTextBoxes.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
