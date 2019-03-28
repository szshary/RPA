using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;

namespace RPA.Core
{
    sealed public class ExcelRuleSet : StatefulRuleSetDecorator
    {
        private Application _excelApplication;
        private Workbook _excelWorkbook;

        public ExcelRuleSet(StatefulRuleSet statefulRuleSet) : base(statefulRuleSet)
        {
            _elementStartRules.Add("ClearExcelRange", ClearExcelRange);
            _elementStartRules.Add("CopyExcelCellToExcelCell", CopyExcelCellToExcelCell);
            _elementStartRules.Add("StoreUsedRangeRowCountToVariable", StoreUsedRangeRowCountToVariable);
            _elementStartRules.Add("CopyExcelCellToVariable", CopyExcelCellToVariable);
            _elementStartRules.Add("StoreExcelRangeToTable", StoreExcelRangeToTable);
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

        private void ClearExcelRange(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("TargetWorksheet") && CheckWorksheetExists(parameters["TargetWorksheet"])
                && (parameters.ContainsKey("StartRow") ^ parameters.ContainsKey("StartRowVariable"))
                && (parameters.ContainsKey("StartColumn") ^ parameters.ContainsKey("StartColumnVariable"))
                && (parameters.ContainsKey("EndRow") ^ parameters.ContainsKey("EndRowVariable"))
                && (parameters.ContainsKey("EndColumn") ^ parameters.ContainsKey("EndColumnVariable")))
            {
                int startRow = 0, startColumn = 0, endRow = 0, endColumn = 0;

                if (parameters.ContainsKey("StartRow"))
                {
                    Int32.TryParse(parameters["StartRow"], out startRow);
                }
                else
                {
                    Int32.TryParse(EngineState.VariableDictionary[parameters["StartRowVariable"]].ToString(), out startRow);
                }
                if (parameters.ContainsKey("StartColumn"))
                {
                    Int32.TryParse(parameters["StartColumn"], out startColumn);
                }
                else
                {
                    Int32.TryParse(EngineState.VariableDictionary[parameters["StartColumnVariable"]].ToString(), out startColumn);
                }
                if (parameters.ContainsKey("EndRow"))
                {
                    Int32.TryParse(parameters["EndRow"], out endRow);
                }
                else
                {
                    Int32.TryParse(EngineState.VariableDictionary[parameters["EndRowVariable"]].ToString(), out endRow);
                }
                if (parameters.ContainsKey("EndColumn"))
                {
                    Int32.TryParse(parameters["EndColumn"], out endColumn);
                }
                else
                {
                    Int32.TryParse(EngineState.VariableDictionary[parameters["EndColumnVariable"]].ToString(), out endColumn);
                }
                if (startRow != 0 && startColumn != 0 && endRow != 0 && endColumn != 0)
                {
                    Worksheet targetSheet = _excelWorkbook.Sheets[parameters["TargetWorksheet"]];
                    targetSheet.Range[targetSheet.Cells[startRow, startColumn], targetSheet.Cells[endRow, endColumn]].Clear();
                }
            }
        }

        private void CopyExcelCellToExcelCell(Dictionary<String, String> parameters)
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
                    Int32.TryParse(EngineState.VariableDictionary[parameters["SourceRowVariable"]].ToString(), out sourceRow);
                }
                if (parameters.ContainsKey("SourceColumn"))
                {
                    Int32.TryParse(parameters["SourceColumn"], out sourceColumn);
                }
                else
                {
                    Int32.TryParse(EngineState.VariableDictionary[parameters["SourceColumnVariable"]].ToString(), out sourceColumn);
                }
                if (parameters.ContainsKey("TargetRow"))
                {
                    Int32.TryParse(parameters["TargetRow"], out targetRow);
                }
                else
                {
                    Int32.TryParse(EngineState.VariableDictionary[parameters["TargetRowVariable"]].ToString(), out targetRow);
                }
                if (parameters.ContainsKey("TargetColumn"))
                {
                    Int32.TryParse(parameters["TargetColumn"], out targetColumn);
                }
                else
                {
                    Int32.TryParse(EngineState.VariableDictionary[parameters["TargetColumnVariable"]].ToString(), out targetColumn);
                }
                if (sourceRow != 0 && sourceColumn != 0 && targetRow != 0 && targetColumn != 0)
                {
                    _excelWorkbook.Sheets[parameters["TargetWorksheet"]].Cells[targetRow, targetColumn].Value = _excelWorkbook.Sheets[parameters["SourceWorksheet"]].Cells[sourceRow, sourceColumn].Value;
                }
            }
        }

        private void StoreUsedRangeRowCountToVariable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Variable") && parameters.ContainsKey("SourceWorksheet") && CheckWorksheetExists(parameters["SourceWorksheet"]))
            {
                Range excelRange = _excelWorkbook.Sheets[parameters["SourceWorksheet"]].UsedRange;

                if (EngineState.VariableDictionary.ContainsKey(parameters["Variable"]))
                {
                    EngineState.VariableDictionary[parameters["Variable"]] = excelRange.Rows.Count;
                }
                else
                {
                    EngineState.VariableDictionary.Add(parameters["Variable"], excelRange.Rows.Count);
                }
            }
        }

        private void CopyExcelCellToVariable(Dictionary<String, String> parameters)
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
                    Int32.TryParse(EngineState.VariableDictionary[parameters["SourceRowVariable"]].ToString(), out sourceRow);
                }
                if (parameters.ContainsKey("SourceColumn"))
                {
                    Int32.TryParse(parameters["SourceColumn"], out sourceColumn);
                }
                else
                {
                    Int32.TryParse(EngineState.VariableDictionary[parameters["SourceColumnVariable"]].ToString(), out sourceColumn);
                }
                if (sourceColumn != 0 && sourceColumn != 0)
                {
                    if (EngineState.VariableDictionary.ContainsKey(parameters["Variable"]))
                    {
                        EngineState.VariableDictionary[parameters["Variable"]] = _excelWorkbook.Sheets[parameters["SourceWorksheet"]].Cells[sourceRow, sourceColumn].Value;
                    }
                    else
                    {
                        EngineState.VariableDictionary.Add(parameters["Variable"], _excelWorkbook.Sheets[parameters["SourceWorksheet"]].Cells[sourceRow, sourceColumn].Value);
                    }
                }
            }
        }

        private void StoreExcelRangeToTable(Dictionary<String, String> parameters)
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
                    if (EngineState.TableList.Exists((x) => { return (x.TableName == table.TableName); }))
                    {
                        EngineState.TableList.Remove(EngineState.TableList.Find((x) => { return (x.TableName == table.TableName); }));
                    }
                    table.TableName = parameters["Table"];
                    EngineState.TableList.Add(table);
                }
            }
        }

        private void WriteToExcelCell(Dictionary<String, String> parameters)
        {
            if ((parameters.ContainsKey("Value") ^ (parameters.ContainsKey("Variable") && EngineState.VariableDictionary.ContainsKey(parameters["Variable"]))
                ^ (parameters.ContainsKey("Table") && EngineState.TableList.Exists((x) => { return (x.TableName == parameters["Table"]); })))
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
                    Int32.TryParse(EngineState.VariableDictionary[parameters["TargetRowVariable"]].ToString(), out targetRow);
                }
                if (parameters.ContainsKey("TargetColumn"))
                {
                    Int32.TryParse(parameters["TargetColumn"], out targetColumn);
                }
                else
                {
                    Int32.TryParse(EngineState.VariableDictionary[parameters["TargetColumnVariable"]].ToString(), out targetColumn);
                }
                if (targetRow != 0 && targetColumn != 0)
                {
                    if (parameters.ContainsKey("Value"))
                    {
                        _excelWorkbook.Sheets[parameters["TargetWorksheet"]].Cells[targetRow, targetColumn].Value = parameters["Value"];
                    }
                    else if (parameters.ContainsKey("Variable"))
                    {
                        _excelWorkbook.Sheets[parameters["TargetWorksheet"]].Cells[targetRow, targetColumn].Value = EngineState.VariableDictionary[parameters["Variable"]];
                    }
                    else
                    {
                        System.Data.DataTable table = EngineState.TableList.Find((x) => { return (x.TableName == parameters["Table"]); });

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            for (int j = 0; j < table.Columns.Count; j++)
                            {
                                _excelWorkbook.Sheets[parameters["TargetWorksheet"]].Cells[targetRow + i, targetColumn + j].Value = table.Rows[i].ItemArray[j];
                            }
                        }
                    }
                }
            }
        }

        private void StartExcelSession(Dictionary<String, String> parameters)
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

        private void EndExcelSession()
        {
            if (_excelApplication != null)
            {
                if (_excelWorkbook != null)
                {
                    _excelWorkbook.Save();
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