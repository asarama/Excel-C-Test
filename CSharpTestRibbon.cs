using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
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
        private int input_worksheet_tenant_insert_row = 3;
        private string[] input_worksheet_tenant_column_search_range = new string[2] { "A", "J" };

        private int market_worksheet_index = 3;
        private int market_worksheet_tenant_insert_row = 9;
        private string[] market_worksheet_tenant_column_search_range = new string[2] { "B", "DL" };
        private int[] market_worksheet_columns = new int[2] { 17, 117 }; //26 is Q and 50 is DL

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

        // DELETE ROW button handler
        // Find location of fist empty row in input and market sheet
        // Delete the row before it
        public void OnDeleteRowButton(Office.IRibbonControl control)
        {
            // Find last row in input worksheet
            string input_worksheet_target_row = FindNextEmptyRow(
                this.input_worksheet_tenant_column_search_range,
                this.input_worksheet_tenant_insert_row,
                this.input_worksheet_index
            );
            int input_worksheet_target_row_number = Int32.Parse(input_worksheet_target_row);

            // Do not allow deleting first row
            if (input_worksheet_target_row_number == 4)
            {
                MessageBox.Show("Can not delete first row!");
                return;
            }

            // Delete the row
            Excel.Worksheet input_worksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets[this.input_worksheet_index]);
            Excel.Range input_worksheet_target_row_range = input_worksheet.get_Range($"A{input_worksheet_target_row_number - 1}");
            input_worksheet_target_row_range.EntireRow.Delete();

            // Find last row in market worksheet
            string market_worksheet_target_row = FindNextEmptyRow(
                this.market_worksheet_tenant_column_search_range,
                this.market_worksheet_tenant_insert_row,
                this.market_worksheet_index
            );
            int market_worksheet_target_row_number = Int32.Parse(market_worksheet_target_row);

            // Delete the row
            Excel.Worksheet market_worksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets[this.market_worksheet_index]);
            Excel.Range market_worksheet_target_row_range = market_worksheet.get_Range($"A{market_worksheet_target_row_number - 1}");
            market_worksheet_target_row_range.EntireRow.Delete();

        }

        // ADD ROW button handler
        public void OnAddRowButton(Office.IRibbonControl control)
        {
            // Add row to input sheet
            int input_row_number = this.AddRowToInputWorksheet();

            // Add row to market assumptions row
            this.AddRowToMarketWorksheet(input_row_number);
        }

        #endregion

        #region Main Logic

        // Find location of first empty row in input sheet
        // Add row with tenant name
        // Tenant name should be Tenant ${row_number - 1}
        private int AddRowToInputWorksheet()
        {
            // Find next available row
            string input_worksheet_target_row = FindNextEmptyRow(
                this.input_worksheet_tenant_column_search_range,
                this.input_worksheet_tenant_insert_row,
                this.input_worksheet_index
            );
            int input_worksheet_target_row_number = Int32.Parse(input_worksheet_target_row);
            string tenant_name = $"Tenant {input_worksheet_target_row_number - 2}";

            // Get row instance and insert row above it
            Excel.Worksheet input_worksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets[this.input_worksheet_index]);
            Excel.Range input_worksheet_empty_row = input_worksheet.get_Range($"A{input_worksheet_target_row}");
            // Since we insert our formating stays consistent
            input_worksheet_empty_row.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            // Insert new data
            SetWorksheetRowValue(input_worksheet, $"A{input_worksheet_target_row}", tenant_name);

            return input_worksheet_target_row_number;
        }

        // Find location of empty row in market sheet
        // Add row with tenant name while adding the expected formulas for the market assumption columns
        private void AddRowToMarketWorksheet(int input_row_number)
        {
            // Find next available row
            string market_worksheet_target_row = FindNextEmptyRow(
                this.market_worksheet_tenant_column_search_range,
                this.market_worksheet_tenant_insert_row,
                this.market_worksheet_index
            );
            int market_worksheet_target_row_number = Int32.Parse(market_worksheet_target_row);
            string market_input_value = $"=Input!C{input_row_number.ToString()}";

            // Get row instance and insert row above it
            Excel.Worksheet market_worksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets[this.market_worksheet_index]);
            Excel.Range market_worksheet_empty_row = market_worksheet.get_Range($"B{market_worksheet_target_row}");
            // Since we insert our formating stays consistent
            market_worksheet_empty_row.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            // Insert new data
            SetWorksheetRowValue(market_worksheet, $"B{market_worksheet_target_row}", $"=Input!A{input_row_number.ToString()}");
            SetWorksheetRowValue(market_worksheet, $"D{market_worksheet_target_row}", market_input_value);
            SetWorksheetRowValue(market_worksheet, $"E{market_worksheet_target_row}:P{market_worksheet_target_row}", "0");
            // Fill in from from rows Q to DL
            int current_column_number = this.market_worksheet_columns[0];
            while (current_column_number < this.market_worksheet_columns[1])
            {
                string market_assumption_equation_value = this.CreateMarketAssumptionCellValue(current_column_number, market_worksheet_target_row_number);
                SetWorksheetRowValue(market_worksheet, this.Integer2ExcelColumn(current_column_number) + market_worksheet_target_row, market_assumption_equation_value);
                current_column_number++;
            }
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
                string cell_value = String.Empty;
                try
                {
                    cell_value = cell.ToString();
                } catch(Exception e)
                {
                    cell_value = ((String)cell);
                }
                if (!String.IsNullOrEmpty(cell_value))
                {
                    return false;
                }
            }

            return true;
        }

        // Update row range value
        private static void SetWorksheetRowValue(Excel.Worksheet target_worksheet, string target_range_string, string update_value)
        {
            Excel.Range target_row = target_worksheet.get_Range(target_range_string);
            target_row.Value2 = update_value;
        }

        private int MAX_CAP_CHAR_NUMBER = 97;
        private int CHAR_IN_ALPHABET = 26;
        // Used to convert numbers to excel columns
        // From: https://stackoverflow.com/questions/181596/how-to-convert-a-column-number-eg-127-into-an-excel-column-eg-aa
        private string Integer2ExcelColumn(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % this.CHAR_IN_ALPHABET;
                columnName = Convert.ToChar(this.MAX_CAP_CHAR_NUMBER + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / this.CHAR_IN_ALPHABET);
            }

            return columnName.ToUpper();

        }

        // Create market assumption cell value
        private string CreateMarketAssumptionCellValue(int column_number, int row_number)
        {
            string column_name = this.Integer2ExcelColumn(column_number);
            string previous_column_name = this.Integer2ExcelColumn(column_number - 1);

            return $"= IF(OR({column_name}$4 < Input!$G{row_number},{column_name}$4 > Input!$J{row_number}),0,IF({column_name}$4 = Input!$G{row_number},Input!$D{row_number} / 12,IF(AVERAGE({this.Integer2ExcelColumn(column_number - 12)}{row_number}:{previous_column_name}{row_number})<>{previous_column_name}{row_number},{previous_column_name}{row_number},IF(Input!$E{row_number} = \"$/SF\",{previous_column_name}{row_number}+Input!$F{row_number},IF(Input!$E{row_number} = \" % \", Q{row_number} * (1 + Input!$F{row_number} / 100))))))";
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
