// Title: How to set and then remove an opening password on an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a new workbook, assigns an opening password, saves it as XLSX, then opens the file using LoadOptions with the password, clears the password, and saves the workbook unchanged. | Provide a step‑by‑step C# example that protects an Excel file with an opening password and later removes that protection without altering any worksheet data using Aspose.Cells.
// Common Searches: aspnet set opening password on Excel file with Aspose.Cells and later remove it | c# remove workbook opening password using Aspose.Cells LoadOptions | how to clear password protection from an XLSX file programmatically with Aspose.Cells | Aspose.Cells example for adding and deleting workbook password without changing content | load password protected workbook and save without password using C# Aspose.Cells
// Tags: Aspose.Cells set workbook opening password C# | Aspose.Cells clear workbook password C# | LoadOptions password protected XLSX Aspose.Cells | Save workbook unchanged after password removal Aspose.Cells | Excel file protection handling Aspose.Cells .NET

using Aspose.Cells;
using System;

// Creates a new workbook, applies an opening password, saves it as XLSX, reloads the file with LoadOptions using the password, clears the password, and saves the workbook again without modifying its content.
class Program
{
    static void Main()
    {
        // Path to the workbook file
        string filePath = "ProtectedWorkbook.xlsx";

        // 1. Create a new workbook and apply an opening password
        Workbook wb = new Workbook();
        wb.Settings.Password = "Open123"; // set opening password
        wb.Save(filePath, SaveFormat.Xlsx); // save the protected workbook

        // 2. Load the workbook using the opening password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.Password = "Open123"; // provide password for opening
        Workbook loadedWb = new Workbook(filePath, loadOptions);

        // 3. Remove the opening password
        loadedWb.Settings.Password = null; // clearing the password removes protection

        // 4. Save the workbook back (content remains unchanged)
        loadedWb.Save(filePath, SaveFormat.Xlsx);
    }
}
