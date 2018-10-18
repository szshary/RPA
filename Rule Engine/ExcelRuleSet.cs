using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;

namespace RPA.Core
{
    sealed public class ExcelRuleSet : RuleSetDecorator
    {
        private Application _excelApplication;
        private Workbook _excelWorkbook;

        public ExcelRuleSet(RuleSet _ruleSet) : base(_ruleSet)
        {
            _elementStartRules.Add("StoreUsedRangeRowCountToVariable", StoreUsedRangeRowCountToVariable);
            _elementStartRules.Add("StoreExcelCellToVariable", StoreExcelCellToVariable);
            _elementStartRules.Add("StoreExcelRangeToTable", StoreExcelRangeToTable);
            _elementStartRules.Add("CopyExcelCellToExcelCell", CopyExcelCellToExcelCell);
            _elementStartRules.Add("WriteToExcelCell", WriteToExcelCell);

            _elementStartRules.Add("ExcelSession", StartExcelSession);
            _elementEndRules.Add("ExcelSession", EndExcelSession);
        }

        private bool CheckWorksheetExists(String sheetName)
        {
            foreach (_Worksheet sheet in _excelWorkbook.Sheets)
            {
                if (sheet.Name == sheetName)
                {
                    return true;
                }
            }
            return false;
        }

        private void StoreUsedRangeRowCountToVariable(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (parameters.ContainsKey("Variable") && parameters.ContainsKey("SourceWorksheet") && CheckWorksheetExists(parameters["SourceWorksheet"]))
            {
                Range excelRange = _excelWorkbook.Sheets[parameters["SourceWorksheet"]].UsedRange;

                if (engineState.VariableCollection.ContainsKey(parameters["Variable"]))
                {
                    engineState.VariableCollection[parameters["Variable"]] = excelRange.Rows.Count;
                }
                else
                {
                    engineState.VariableCollection.Add(parameters["Variable"], excelRange.Rows.Count);
                }
            }
        }

        private void StoreExcelCellToVariable(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (parameters.ContainsKey("Variable") && parameters.ContainsKey("SourceWorksheet") && CheckWorksheetExists(parameters["SourceWorksheet"]))
            {
                int sourceRow = 0, sourceColumn = 0;

                if (parameters.ContainsKey("SourceRow"))
                {
                    Int32.TryParse(parameters["SourceRow"], out sourceRow);
                }
                else
                {
                    Int32.TryParse(engineState.VariableCollection[parameters["SourceRowVariable"]].ToString(), out sourceRow);
                }
                if (parameters.ContainsKey("SourceColumn"))
                {
                    Int32.TryParse(parameters["SourceColumn"], out sourceColumn);
                }
                else
                {
                    Int32.TryParse(engineState.VariableCollection[parameters["SourceColumnVariable"]].ToString(), out sourceColumn);
                }
                if (sourceColumn != 0 && sourceColumn != 0)
                {
                    if (engineState.VariableCollection.ContainsKey(parameters["Variable"]))
                    {
                        engineState.VariableCollection[parameters["Variable"]] = _excelWorkbook.Sheets[parameters["SourceWorksheet"]].Cells[sourceRow, sourceColumn].Value;
                    }
                    else
                    {
                        engineState.VariableCollection.Add(parameters["Variable"], _excelWorkbook.Sheets[parameters["SourceWorksheet"]].Cells[sourceRow, sourceColumn].Value);
                    }
                }
            }
        }

        private void StoreExcelRangeToTable(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (parameters.ContainsKey("Table") && parameters.ContainsKey("SourceWorksheet") && CheckWorksheetExists(parameters["SourceWorksheet"]))
            {
                Range excelRange = _excelWorkbook.Sheets[parameters["SourceWorksheet"]].UsedRange;

                using (System.Data.DataTable table = new System.Data.DataTable(parameters["Table"]))
                {
                    Object[,] values = (Object[,])excelRange.Value2;

                    for (int i = 1; i <= values.GetLength(1); i++)
                    {
                        table.Columns.Add(values[1, i].ToString());
                    }
                    for (int i = 2; i <= values.GetLength(0); ++i)
                    {
                        DataRow row = table.NewRow();
                        for (var j = 1; j <= values.GetLength(1); ++j)
                        {
                            row[j - 1] = values[i, j];
                        }
                        table.Rows.Add(row);
                    }
                    if (engineState.TableCollection.Exists((x) => { return (x.TableName == table.TableName); }))
                    {
                        engineState.TableCollection.Remove(engineState.TableCollection.Find((x) => { return (x.TableName == table.TableName); }));
                    }
                    table.TableName = parameters["Table"];
                    engineState.TableCollection.Add(table);
                }
            }
        }

        private void CopyExcelCellToExcelCell(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (parameters.ContainsKey("SourceWorksheet") && parameters.ContainsKey("TargetWorksheet")
                && (parameters.ContainsKey("SourceRow") ^ parameters.ContainsKey("SourceRowVariable"))
                && (parameters.ContainsKey("SourceColumn") ^ parameters.ContainsKey("SourceColumnVariable"))
                && (parameters.ContainsKey("TargetRow") ^ parameters.ContainsKey("TargetRowVariable"))
                && (parameters.ContainsKey("TargetColumn") ^ parameters.ContainsKey("TargetColumnVariable")))
            {
                int sourceRow = 0, sourceColumn = 0, targetRow = 0, targetColumn = 0;

                if (parameters.ContainsKey("SourceRow"))
                {
                    Int32.TryParse(parameters["SourceRow"], out sourceRow);
                }
                else
                {
                    Int32.TryParse(engineState.VariableCollection[parameters["SourceRowVariable"]].ToString(), out sourceRow);
                }
                if (parameters.ContainsKey("SourceColumn"))
                {
                    Int32.TryParse(parameters["SourceColumn"], out sourceColumn);
                }
                else
                {
                    Int32.TryParse(engineState.VariableCollection[parameters["SourceColumnVariable"]].ToString(), out sourceColumn);
                }
                if (parameters.ContainsKey("TargetRow"))
                {
                    Int32.TryParse(parameters["TargetRow"], out targetRow);
                }
                else
                {
                    Int32.TryParse(engineState.VariableCollection[parameters["TargetRowVariable"]].ToString(), out targetRow);
                }
                if (parameters.ContainsKey("TargetColumn"))
                {
                    Int32.TryParse(parameters["TargetColumn"], out targetColumn);
                }
                else
                {
                    Int32.TryParse(engineState.VariableCollection[parameters["TargetColumnVariable"]].ToString(), out targetColumn);
                }
                if (sourceRow != 0 && sourceColumn != 0 && targetRow != 0 && targetColumn != 0)
                {
                    _excelWorkbook.Sheets[parameters["TargetWorksheet"]].Cells[targetRow, targetColumn].Value = _excelWorkbook.Sheets[parameters["SourceWorksheet"]].Cells[sourceRow, sourceColumn].Value;
                }
            }
        }

        private void WriteToExcelCell(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if ((parameters.ContainsKey("Value") ^ (parameters.ContainsKey("Variable") && engineState.VariableCollection.ContainsKey(parameters["Variable"])))
                && parameters.ContainsKey("TargetWorksheet") && CheckWorksheetExists(parameters["TargetWorksheet"])
                && (parameters.ContainsKey("TargetRow") ^ parameters.ContainsKey("TargetRowVariable"))
                && (parameters.ContainsKey("TargetColumn") ^ parameters.ContainsKey("TargetColumnVariable")))
            {
                int targetRow = 0, targetColumn = 0;

                if (parameters.ContainsKey("TargetRow"))
                {
                    Int32.TryParse(parameters["TargetRow"], out targetRow);
                }
                else
                {
                    Int32.TryParse(engineState.VariableCollection[parameters["TargetRowVariable"]].ToString(), out targetRow);
                }
                if (parameters.ContainsKey("TargetColumn"))
                {
                    Int32.TryParse(parameters["TargetColumn"], out targetColumn);
                }
                else
                {
                    Int32.TryParse(engineState.VariableCollection[parameters["TargetColumnVariable"]].ToString(), out targetColumn);
                }
                if (targetRow != 0 && targetColumn != 0)
                {
                    if (parameters.ContainsKey("Value"))
                    {
                        _excelWorkbook.Sheets[parameters["TargetWorksheet"]].Cells[targetRow, targetColumn].Value = parameters["Value"];
                    }
                    else
                    {
                        _excelWorkbook.Sheets[parameters["TargetWorksheet"]].Cells[targetRow, targetColumn].Value = engineState.VariableCollection[parameters["Variable"]];
                    }
                }
            }
        }

        private void StartExcelSession(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            Process[] orphanExcelProcesses = Process.GetProcessesByName("EXCEL");
            foreach (Process orphanProcess in orphanExcelProcesses)
            {
                if (orphanProcess.MainWindowTitle.Length == 0)
                {
                    orphanProcess.Kill();
                }
            }
            if (_excelApplication == null)
            {
                _excelApplication = new Application
                {
                    Visible = true,
                    WindowState = XlWindowState.xlMaximized
                };
            }
            if (parameters.ContainsKey("FileName"))
            {
                if (File.Exists(Path.Combine(Directory.GetCurrentDirectory(), (parameters.ContainsKey("Folder") ? parameters["Folder"] : String.Empty), parameters["FileName"])))
                {
                    _excelWorkbook = _excelApplication.Workbooks.Open(Path.Combine(Directory.GetCurrentDirectory(), (parameters.ContainsKey("Folder") ? parameters["Folder"] : String.Empty), parameters["FileName"]));
                }
            }
        }

        private void EndExcelSession(RuleEngineState engineState)
        {
            if (_excelApplication != null)
            {
                if (_excelWorkbook != null)
                {
                    _excelWorkbook.Close();
                }
                _excelApplication.WindowState = XlWindowState.xlMinimized;
                _excelApplication.Quit();

                Process[] orphanExcelProcesses = Process.GetProcessesByName("EXCEL");
                foreach (Process orphanProcess in orphanExcelProcesses)
                {
                    if (orphanProcess.MainWindowTitle.Length == 0)
                    {
                        orphanProcess.Kill();
                    }
                }
            }
        }
    }
}