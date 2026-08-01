// Title: Set separate open and edit passwords on an XLS workbook using Aspose.Cells for .NET
// Description: Shows how to create an XLS file, assign a password needed to open it and a different one needed to edit, save the workbook, then load it with the open credential and confirm edit protection via the Aspose.Cells C# API.
// Keywords: Aspose.Cells | C# | .NET | XLS encryption | open password | edit password | write protection | Workbook.Settings.Password | WriteProtection.Password | password protected Excel | LoadOptions password
// Common Searches: Aspose.Cells set open password C# | how to add edit password to XLS with Aspose.Cells | load password protected workbook .NET | validate write protection password Aspose.Cells | difference between open and edit passwords in Aspose.Cells
// Developer Intent: Implement distinct credentials for opening and modifying an XLS file and programmatically verify both safeguards.
// Use Cases: Distribute read‑only Excel reports that require a view password, while allowing authorized users to edit with a second credential. | Generate templates that anyone can open but only users with a separate password can modify. | Add password validation before automatically updating a protected workbook in a .NET automation pipeline.
// AI Prompts: Generate C# code to change the edit password of an existing XLS workbook using Aspose.Cells. | Show how to remove write protection from a workbook after providing the correct edit password. | Explain steps to protect a .xlsx file with an open password and a separate edit password using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to create an XLS file, assign a password needed to open it and a different one needed to edit, save the workbook, then load it with the open credential and confirm edit protection via the Aspose.Cells C# API.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        workbook.Worksheets[0].Cells["A1"].PutValue("Sensitive data");

        // ----- Set opening (file encryption) password -----
        // This password is required to open the workbook
        workbook.Settings.Password = "OpenPass123";

        // ----- Set modifying (write protection) password -----
        // This password is required to modify the workbook after it is opened
        workbook.Settings.WriteProtection.Password = "ModifyPass456";

        // Save the workbook to a file (XLS format)
        string filePath = "ProtectedWorkbook.xls";
        workbook.Save(filePath);

        // ----- Load the workbook using the opening password -----
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "OpenPass123";
        Workbook loadedWorkbook = new Workbook(filePath, loadOptions);

        // Verify that the opening password worked
        Console.WriteLine("Workbook opened successfully with opening password.");

        // Verify that write protection is enabled
        bool isWriteProtected = loadedWorkbook.Settings.WriteProtection.IsWriteProtected;
        Console.WriteLine("Write protection enabled: " + isWriteProtected);

        // Validate the modify password
        bool isModifyPasswordValid = loadedWorkbook.Settings.WriteProtection.ValidatePassword("ModifyPass456");
        Console.WriteLine("Modify password validation result: " + isModifyPasswordValid);
    }
}
