using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaDemo
{
    public class VbaOperations
    {
        // Path for the macro-enabled template (XLTM)
        private const string TemplatePath = "MacroTemplate.xltm";

        // Creates a new XLTM workbook, adds a procedural module with sample code, and saves it.
        public static void CreateWorkbookWithVba()
        {
            // Create a new workbook (default is XLSX)
            Workbook workbook = new Workbook();

            // Ensure a VBA project exists by saving as macro-enabled template and reloading
            workbook.Save(TemplatePath, SaveFormat.Xltm);
            workbook = new Workbook(TemplatePath);

            // Add a new procedural module named "SampleModule"
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Procedural, "SampleModule");
            VbaModule module = workbook.VbaProject.Modules[moduleIndex];

            // Set initial VBA code
            module.Codes = "Sub HelloWorld()\n    MsgBox \"Hello from VBA!\"\nEnd Sub";

            // Save the workbook with the VBA module
            workbook.Save(TemplatePath, SaveFormat.Xltm);
        }

        // Loads the XLTM workbook, modifies the code of an existing module, and saves the changes.
        public static void EditVbaModule()
        {
            // Load the existing macro-enabled template
            Workbook workbook = new Workbook(TemplatePath);

            // Access the module by name (or index)
            VbaModule module = workbook.VbaProject.Modules["SampleModule"];

            // Append a new subroutine to the existing code
            string additionalCode = "\nSub ShowDate()\n    MsgBox \"Today is \" & Date\nEnd Sub";
            module.Codes = module.Codes + additionalCode;

            // Save the updated workbook
            workbook.Save(TemplatePath, SaveFormat.Xltm);
        }

        // Loads the XLTM workbook and extracts the VBA code from a specified module.
        public static void ExtractVbaModule()
        {
            // Load the workbook
            Workbook workbook = new Workbook(TemplatePath);

            // Retrieve the module (by name)
            VbaModule module = workbook.VbaProject.Modules["SampleModule"];

            // Output the VBA code to console (or any other destination)
            Console.WriteLine("=== VBA Code in module 'SampleModule' ===");
            Console.WriteLine(module.Codes);
        }

        // Removes a VBA module from the workbook and saves the result.
        public static void RemoveVbaModule()
        {
            // Load the workbook
            Workbook workbook = new Workbook(TemplatePath);

            // Remove the module by name
            workbook.VbaProject.Modules.Remove("SampleModule");

            // Save the workbook without the removed module
            workbook.Save(TemplatePath, SaveFormat.Xltm);
        }

        // Demonstrates the full workflow.
        public static void RunDemo()
        {
            // Ensure any previous file is cleared
            if (File.Exists(TemplatePath))
                File.Delete(TemplatePath);

            // 1. Create workbook with VBA
            CreateWorkbookWithVba();

            // 2. Edit the existing VBA module
            EditVbaModule();

            // 3. Extract and display the VBA code
            ExtractVbaModule();

            // 4. Remove the VBA module
            RemoveVbaModule();

            // 5. Verify removal by attempting to extract again (will throw if not found)
            try
            {
                ExtractVbaModule();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Module extraction after removal failed as expected: " + ex.Message);
            }
        }
    }

    // Entry point
    class Program
    {
        static void Main()
        {
            VbaOperations.RunDemo();
        }
    }
}