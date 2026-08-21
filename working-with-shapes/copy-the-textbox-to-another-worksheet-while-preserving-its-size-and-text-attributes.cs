// Title: Copy a TextBox Shape Between Worksheets While Preserving Size and Formatting – Aspose.Cells for .NET
// Description: This C# example shows how to duplicate a TextBox from a source worksheet to a destination worksheet using Aspose.Cells. The Shapes.AddCopy method retains the original row, column and pixel offsets. After copying, the code casts the shape back to TextBox and transfers the text and font properties (name, size, bold) so the appearance is identical. The workbook is then saved as an XLSX file.
// Keywords: Aspose.Cells copy TextBox | duplicate shape worksheet .NET | preserve shape size Aspose.Cells | AddCopy shape method | copy TextBox formatting C# | Aspose.Cells shape cloning | C# Excel textbox copy | Aspose.Cells workbook template
// Common Searches: how to copy a TextBox between worksheets using Aspose.Cells | Aspose.Cells preserve textbox size and font | C# copy shape with formatting Aspose.Cells | AddCopy example Aspose.Cells .NET | duplicate Excel textbox programmatically
// Developer Intent: Duplicate a TextBox from one worksheet to another while keeping its position, dimensions, and font styling unchanged.
// Use Cases: Reuse a styled TextBox header across multiple report sheets for consistent branding. | Create a dashboard where the same annotation box appears on each worksheet without manual recreation. | Generate template‑based workbooks that require the same pre‑formatted TextBox on every new sheet.
// AI Prompts: Generate a reusable C# method that copies any shape (TextBox, rectangle, etc.) and returns the new shape with all visual properties preserved. | Show how to copy a TextBox to several worksheets in a loop, maintaining exact position and style on each sheet. | Explain how to copy a TextBox along with its hyperlink and fill color using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This C# example shows how to duplicate a TextBox from a source worksheet to a destination worksheet using Aspose.Cells. The Shapes.AddCopy method retains the original row, column and pixel offsets. After copying, the code casts the shape back to TextBox and transfers the text and font properties (name, size, bold) so the appearance is identical. The workbook is then saved as an XLSX file.
    public class CopyTextBoxDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // ---------- Source worksheet ----------
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Source";

                // Add a TextBox shape to the source sheet
                // Parameters: upper left row, top offset (pixels), upper left column, left offset (pixels), width (pixels), height (pixels)
                TextBox sourceTextBox = (TextBox)sourceSheet.Shapes.AddTextBox(2, 0, 2, 0, 200, 80);
                sourceTextBox.Text = "Hello Aspose!";
                sourceTextBox.Font.Name = "Calibri";
                sourceTextBox.Font.Size = 14;
                sourceTextBox.Font.IsBold = true;

                // ---------- Destination worksheet ----------
                Worksheet destSheet = workbook.Worksheets.Add("Destination");

                // Copy the TextBox to the destination sheet preserving its position and size
                // AddCopy copies the shape and retains its dimensions
                Shape copiedShape = destSheet.Shapes.AddCopy(
                    sourceTextBox,                     // source shape
                    sourceTextBox.UpperLeftRow,        // top row index
                    sourceTextBox.Y,                   // vertical offset (pixels)
                    sourceTextBox.UpperLeftColumn,     // left column index
                    sourceTextBox.X);                  // horizontal offset (pixels)

                // Cast the copied shape back to TextBox to transfer text attributes
                TextBox destTextBox = (TextBox)copiedShape;
                destTextBox.Text = sourceTextBox.Text;
                destTextBox.Font.Name = sourceTextBox.Font.Name;
                destTextBox.Font.Size = sourceTextBox.Font.Size;
                destTextBox.Font.IsBold = sourceTextBox.Font.IsBold;

                // Save the workbook
                workbook.Save("CopyTextbox.xlsx");
                Console.WriteLine("Workbook saved as CopyTextbox.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CopyTextBoxDemo.Run();
        }
    }
}
