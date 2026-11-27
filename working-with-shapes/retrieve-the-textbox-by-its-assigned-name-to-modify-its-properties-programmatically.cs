// Title: Retrieve a TextBox by its assigned name and modify its properties with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find textbox by its Name property in a worksheet | how to get a specific TextBox object from an Aspose.Cells workbook | programmatically change font color and fill of a named textbox using Aspose.Cells | iterate through TextBox collection to locate a shape by name in Aspose.Cells | set textbox placement to MoveAndSize after retrieving it in Aspose.Cells .NET
// Tags: Aspose.Cells named textbox retrieval | C# modify Aspose.Cells textbox appearance | Aspose.Cells textbox placement MoveAndSize | iterate Aspose.Cells TextBox collection | Aspose.Cells shape property update C#

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextBoxDemo
{
    // The example creates a workbook, adds a TextBox with a unique Name, iterates the worksheet's TextBox collection to locate that shape, updates its text, font, fill, line style, and placement, and saves the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Name = "DemoSheet";

                // -----------------------------------------------------------------
                // 2. Add a textbox to the worksheet.
                //    The Add method returns the index of the newly created textbox.
                // -----------------------------------------------------------------
                // Parameters: upper left row, upper left column, height (pixels), width (pixels)
                int textboxIndex = worksheet.TextBoxes.Add(2, 2, 120, 250);
                TextBox addedTextBox = worksheet.TextBoxes[textboxIndex];

                // Assign a unique name to the textbox so we can retrieve it later.
                addedTextBox.Name = "MySpecialTextBox";

                // Set some initial properties.
                addedTextBox.Text = "Initial Text";
                addedTextBox.Font.Color = Color.Blue;
                addedTextBox.Font.Size = 12;
                addedTextBox.Font.IsBold = true;
                addedTextBox.Fill.SolidFill.Color = Color.LightYellow;
                addedTextBox.Line.Weight = 2;
                // Line color may not be supported in all versions; omitted.

                // -----------------------------------------------------------------
                // 3. Retrieve the textbox by its assigned name.
                //    TextBoxCollection does not provide a direct GetByName method,
                //    so we iterate through the collection.
                // -----------------------------------------------------------------
                TextBox targetTextBox = null;
                foreach (TextBox tb in worksheet.TextBoxes)
                {
                    if (tb.Name == "MySpecialTextBox")
                    {
                        targetTextBox = tb;
                        break;
                    }
                }

                if (targetTextBox == null)
                {
                    Console.WriteLine("Textbox with the specified name was not found.");
                    return;
                }

                // -----------------------------------------------------------------
                // 4. Modify the retrieved textbox's properties programmatically.
                // -----------------------------------------------------------------
                targetTextBox.Text = "Updated Text after retrieval";
                targetTextBox.Font.Color = Color.Green;
                targetTextBox.Font.IsItalic = true;
                targetTextBox.Font.Size = 14;
                targetTextBox.Fill.SolidFill.Color = Color.LightCyan;
                targetTextBox.Line.DashStyle = MsoLineDashStyle.Solid;
                targetTextBox.Line.Weight = 3;
                // Line color property may not be available; omitted.

                // Optionally, change placement so the textbox moves with cells.
                targetTextBox.Placement = PlacementType.MoveAndSize;

                // -----------------------------------------------------------------
                // 5. Save the workbook to verify the changes.
                // -----------------------------------------------------------------
                string outputPath = "TextBoxByNameDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
