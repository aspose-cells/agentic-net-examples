// Title: Add a password‑protected workbook and embed the password in a textbox shape using Aspose.Cells for .NET (C#)
// Common Searches: how to protect an entire workbook with a password using Aspose.Cells C# | insert a textbox shape with custom text into an Excel worksheet using Aspose.Cells | save a password‑protected .xlsx file with Aspose.Cells .NET API | display workbook protection password inside a worksheet shape in C# | Aspose.Cells example for protecting a workbook and adding a shape
// Tags: Aspose.Cells workbook protection password C# | Aspose.Cells add textbox shape to worksheet | Aspose.Cells save protected xlsx file | Aspose.Cells shape text formatting C# | Aspose.Cells protect entire workbook

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, applies full protection with a specified password, inserts a textbox shape containing the password text on the first worksheet, and saves the file as a password‑protected .xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the protection password
            string password = "SecurePass123";

            // Protect the entire workbook with the password
            workbook.Protect(ProtectionType.All, password);

            // Add a text box shape that contains the password (used as a QR code placeholder)
            // Position: row 1, column 1; Offsets: 0,0; Size: 150x50 points
            TextBox passwordShape = sheet.Shapes.AddTextBox(1, 1, 0, 0, 150, 50);
            passwordShape.Text = password;
            passwordShape.Font.Size = 12;
            passwordShape.Font.Color = Color.Black;

            // Save the workbook to a file
            string outputPath = "ProtectedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
