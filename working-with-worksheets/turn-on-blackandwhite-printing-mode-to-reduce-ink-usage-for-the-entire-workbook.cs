using System;
using Aspose.Cells;

public class SetBlackAndWhitePrinting
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable black‑and‑white printing for every worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.PageSetup.BlackAndWhite = true;
        }

        // Save the workbook with the updated setting
        workbook.Save("BlackAndWhiteWorkbook.xlsx");
    }
}