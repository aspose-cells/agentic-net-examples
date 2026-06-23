using System;
using Aspose.Cells;

class ProtectWorkbookStructure
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Cells["A1"].PutValue("Editable cell");

        // Add a second worksheet
        int secondIndex = workbook.Worksheets.Add();
        Worksheet sheet2 = workbook.Worksheets[secondIndex];
        sheet2.Name = "SecondSheet";
        sheet2.Cells["B2"].PutValue("Another editable cell");

        // Protect only the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "myPassword");

        // Optional: display protection status
        Console.WriteLine("Workbook structure protected: " + workbook.Settings.IsProtected);
        Console.WriteLine("Password set: " + !string.IsNullOrEmpty(workbook.Settings.Password));

        // Save the workbook
        workbook.Save("ProtectedStructure.xlsx");
    }
}