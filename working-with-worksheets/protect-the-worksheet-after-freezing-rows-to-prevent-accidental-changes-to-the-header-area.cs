// Title: Freeze the top row and protect only the header cells in an Aspose.Cells .NET workbook
// AI Prompts: Generate C# code with Aspose.Cells that freezes the first worksheet row, unlocks all cells, locks only the header row, and then applies worksheet protection. | Show how to use the default style to unlock cells, create a locked style for the header, call FreezePanes, and protect the sheet with ProtectionType.All.
// Common Searches: Aspose.Cells C# freeze first row and protect header only | How to lock header row while allowing other cells to be edited in Aspose.Cells | C# Aspose.Cells protect worksheet after FreezePanes | Set worksheet protection with locked header cells using Aspose.Cells .NET | Unlock all cells then lock specific row in Aspose.Cells workbook
// Tags: freeze panes header row Aspose.Cells | lock header cells worksheet protection Aspose.Cells | unlock all cells default style Aspose.Cells | apply locked style to specific row Aspose.Cells | protect worksheet with ProtectionType.All Aspose.Cells | C# Aspose.Cells workbook header protection

using Aspose.Cells;
using System;

// Creates a new workbook, adds a header row, freezes the first row, unlocks all cells via the default style, locks the header cells with a custom style, protects the worksheet using ProtectionType.All, and saves the file as ProtectedSheet.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Amount");

            // Freeze the first row (rows above row 2)
            int totalRows = sheet.Cells.MaxDataRow + 1;
            int totalColumns = sheet.Cells.MaxDataColumn + 1;
            sheet.FreezePanes(1, 0, totalRows, totalColumns);

            // Unlock all cells by modifying the workbook's default style
            Style defaultStyle = workbook.DefaultStyle;
            defaultStyle.IsLocked = false;
            workbook.DefaultStyle = defaultStyle;

            // Create a style for the header row and lock it
            Style headerStyle = workbook.CreateStyle();
            headerStyle.IsLocked = true;

            // Apply the header style to the first row (columns A‑C)
            for (int col = 0; col < 3; col++)
            {
                sheet.Cells[0, col].SetStyle(headerStyle);
            }

            // Protect the worksheet; only unlocked cells are editable
            sheet.Protect(ProtectionType.All);

            // Save the workbook
            workbook.Save("ProtectedSheet.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
