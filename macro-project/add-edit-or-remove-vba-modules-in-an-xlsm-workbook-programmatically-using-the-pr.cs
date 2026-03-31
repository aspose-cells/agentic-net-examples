using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

public class VbaModuleOperations
{
    public static void Run()
    {
        // Create a new workbook (will be saved as macro-enabled later)
        Workbook workbook = new Workbook();

        // ---------- Add a procedural module ----------
        // Add a procedural module named "Utility"
        int procIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "Utility");
        // Retrieve the added module
        VbaModule procModule = workbook.VbaProject.Modules[procIndex];
        // Set VBA code for the module
        procModule.Codes = "Sub HelloWorld()\n    MsgBox \"Hello from VBA!\"\nEnd Sub";

        // ---------- Add a worksheet‑specific module ----------
        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        // Add a module associated with this worksheet
        int sheetModuleIndex = workbook.VbaProject.Modules.Add(sheet);
        VbaModule sheetModule = workbook.VbaProject.Modules[sheetModuleIndex];
        // Optionally set the module name to the worksheet's code name
        sheetModule.Name = sheet.CodeName;
        // Add VBA code that responds to a worksheet event
        sheetModule.Codes = "Private Sub Worksheet_SelectionChange(ByVal Target As Range)\n    MsgBox \"Selection changed\"\nEnd Sub";

        // Save the workbook with the newly added modules
        workbook.Save("AddedModules.xlsm", SaveFormat.Xlsm);

        // ---------- Load, edit, and remove modules ----------
        // Load the saved workbook
        Workbook loaded = new Workbook("AddedModules.xlsm");

        // Edit the code of the procedural module "Utility"
        VbaModule editModule = loaded.VbaProject.Modules["Utility"];
        editModule.Codes = "Sub HelloWorld()\n    MsgBox \"Edited VBA code\"\nEnd Sub";

        // Remove the worksheet‑specific module by its name (the sheet's code name)
        loaded.VbaProject.Modules.Remove(sheet.CodeName);

        // Save the workbook after editing and removal
        loaded.Save("EditedAndRemovedModules.xlsm", SaveFormat.Xlsm);
    }
}

public class Program
{
    public static void Main()
    {
        VbaModuleOperations.Run();
    }
}