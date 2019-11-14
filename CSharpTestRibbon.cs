using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace FirstExcelAddIn
{
    [ComVisible(true)]
    public class CSharpTestRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private int input_worksheet_index = 2;
        private int market_worksheet_index = 3;

        public CSharpTestRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("FirstExcelAddIn.CSharpTestRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        // TODO: Delete row
        // Find location of fist empty row in input sheet
        // Delete the row before it
        // Find corresponding row in Market Assumptions Tab (search by tenant name) and delete it

        public void OnDeleteRowButton(Office.IRibbonControl control)
        {
            // Not implemented            
        }

        // Add row
        // Find location of first empty row in input sheet
        // Tenant name should be Tenant ${row_number - 1}
        // TODO: Add row to Market Assumptions sheet

        public void OnAddRowButton(Office.IRibbonControl control)
        {
            // Find next available row
            string[] column_search_range = new string[2] { "A", "J" };
            string target_row = FindNextEmptyRow(column_search_range, 3, this.input_worksheet_index);

            // Get row instance and insert row above it
            Excel.Worksheet input_worksheet = ((Excel.Worksheet) Globals.ThisAddIn.Application.Worksheets[this.input_worksheet_index]);
            Excel.Range emptyRow = input_worksheet.get_Range($"A{target_row}");
            // Since we insert our formating stays consistent
            emptyRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            // Insert new data
            Excel.Range endRow = input_worksheet.get_Range($"A{target_row}");
            endRow.Value2 = $"Tenant {(Int32.Parse(target_row) - 1).ToString()}";
        }

        #endregion

        #region Helpers

        // Find next empty row given a starting row and column range
        private static string FindNextEmptyRow(string[] column_search_range, int row_offset = 3, int worksheet_index = -1)
        {
            // Get worksheet
            Excel.Worksheet targetWorksheet;
            if (worksheet_index == -1)
            {
                targetWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            } else
            {
                targetWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets[worksheet_index]);
            }

            string current_range;
            string empty_row = null;
            int row_offset_iterator = row_offset;
            
            // Look for next empty row
            while (String.IsNullOrEmpty(empty_row))
            {
                // Get value of cells in provided range at current row iterator value
                current_range = CreateRowRangeString(column_search_range, row_offset_iterator);
                Excel.Range potential_row = targetWorksheet.get_Range(current_range);
                Object[,] cell_data = ((Object[,])potential_row.Value2);

                // Test cell range for empty data
                if (CellRangeIsEmpty(cell_data))
                {
                    empty_row = row_offset_iterator.ToString();
                }

                row_offset_iterator++;
            }

            return empty_row;
        }

        // Find create row range string given search_string and column
        private static string CreateRowRangeString(string[] column_range, int row_offset)
        {
            return $"{column_range[0] + row_offset}:{column_range[1] + row_offset}";
        }

        // Check to see if all cells in given set are empty
        private static bool CellRangeIsEmpty(Object[,] cell_range_values)
        {
            foreach (var cell in cell_range_values)
            {
                // TODO: This check migth fail for cells that have data that can not be cast to a string
                string cell_value = ((string)cell);
                if (!String.IsNullOrEmpty(cell_value))
                {
                    return false;
                }
            }

            return true;
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
