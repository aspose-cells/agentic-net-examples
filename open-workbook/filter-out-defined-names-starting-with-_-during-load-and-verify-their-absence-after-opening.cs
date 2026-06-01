using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsDefinedNameFilter
{
    class Program
    {
        static void Main()
        {
            // Step 1: Create a workbook and add defined names, some start with '_' 
            Workbook wbCreate = new Workbook();
            Worksheet ws = wbCreate.Worksheets[0];

            // Add normal names
            int idx1 = wbCreate.Worksheets.Names.Add("NormalRange1");
            wbCreate.Worksheets.Names[idx1].RefersTo = "=Sheet1!$A$1:$A$5";

            int idx2 = wbCreate.Worksheets.Names.Add("NormalRange2");
            wbCreate.Worksheets.Names[idx2].RefersTo = "=Sheet1!$B$1:$B$5";

            // Add names that start with '_' (to be filtered out)
            int idxHidden1 = wbCreate.Worksheets.Names.Add("_HiddenRange1");
            wbCreate.Worksheets.Names[idxHidden1].RefersTo = "=Sheet1!$C$1:$C$5";

            int idxHidden2 = wbCreate.Worksheets.Names.Add("_HiddenRange2");
            wbCreate.Worksheets.Names[idxHidden2].RefersTo = "=Sheet1!$D$1:$D$5";

            // Save the workbook to a temporary file
            string filePath = "NamesFilterDemo.xlsx";
            wbCreate.Save(filePath);
            wbCreate.Dispose();

            // Step 2: Load the workbook with a LoadFilter (using the provided constructor)
            LoadFilter loadFilter = new LoadFilter();                     // uses LoadFilter() rule
            LoadOptions loadOptions = new LoadOptions();                 // default options
            loadOptions.LoadFilter = loadFilter;                         // assign the filter (rule)

            Workbook wbLoad = new Workbook(filePath, loadOptions);       // load with options (rule)

            // Step 3: Identify defined names that start with '_' and remove them
            NameCollection names = wbLoad.Worksheets.Names;
            List<string> namesToRemove = new List<string>();

            foreach (Name name in names)
            {
                if (name.Text.StartsWith("_"))
                {
                    namesToRemove.Add(name.Text);
                }
            }

            // Remove the collected names using the provided Remove(string[]) method
            if (namesToRemove.Count > 0)
            {
                names.Remove(namesToRemove.ToArray());                  // uses Remove(string[]) rule
            }

            // Step 4: Verify that no defined name starts with '_' after removal
            bool anyUnderscoreNames = false;
            foreach (Name name in names)
            {
                if (name.Text.StartsWith("_"))
                {
                    anyUnderscoreNames = true;
                    break;
                }
            }

            Console.WriteLine("Underscore-prefixed names present after filtering: " + anyUnderscoreNames);
            // Expected output: false

            // Optional: Save the cleaned workbook
            wbLoad.Save("NamesFilterDemo_Cleaned.xlsx");
            wbLoad.Dispose();
        }
    }
}