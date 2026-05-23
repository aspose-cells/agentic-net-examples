using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace InvoiceNumberDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a label shape (row, column, height, width, top, left)
                // In recent Aspose.Cells versions AddLabel returns a Label object directly
                Label label = worksheet.Shapes.AddLabel(0, 0, 30, 200, 0, 0);
                label.Text = "Invoice Numbers:";
                label.Font.Size = 12;
                label.Font.IsBold = true;

                // Add a textbox that will contain the auto‑numbered list
                int textBoxIdx = worksheet.TextBoxes.Add(2, 0, 200, 150);
                TextBox textBox = worksheet.TextBoxes[textBoxIdx];

                // Prepare the text with a placeholder for each line.
                // The bullet will generate the numeric part, and we prepend the custom prefix.
                textBox.Text = "INV-\nINV-\nINV-\nINV-\nINV-";

                // Configure the first paragraph to use auto‑numbered bullets
                TextParagraph paragraph = textBox.TextBody.TextParagraphs[0];
                paragraph.Bullet.Type = BulletType.AutoNumbered;

                // Set the auto‑numbering scheme and starting index
                AutoNumberedBulletValue bullet = (AutoNumberedBulletValue)paragraph.Bullet.BulletValue;
                bullet.AutonumberScheme = TextAutonumberScheme.ArabicPlain; // 1, 2, 3, …
                bullet.StartAt = 1; // Start from 1

                // Ensure the output directory exists
                string outputPath = "InvoiceNumbers.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}