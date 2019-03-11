using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace RPA.Core
{
    public class DefaultRuleSet : StatefulRuleSet
    {
        public DefaultRuleSet() : base()
        {
            _elementStartRules.Add("AddColumnToTable", AddColumnToTable);
            _elementStartRules.Add("ClearTable", ClearTable);
            _elementStartRules.Add("CreateTable", CreateTable);
            _elementStartRules.Add("ExtractVariablesFromTable", ExtractVariablesFromTable);
            _elementStartRules.Add("LogEvent", LogEvent);
            _elementStartRules.Add("LookUpTable", LookUpTable);
            _elementStartRules.Add("MoveVariablesToTable", MoveVariablesToTable);
            _elementStartRules.Add("RemoveAllVariables", RemoveAllVariables);
            _elementStartRules.Add("RemoveColumnFromTable", RemoveColumnFromTable);
            _elementStartRules.Add("RemoveTable", RemoveTable);
            _elementStartRules.Add("RemoveVariable", RemoveVariable);
            _elementStartRules.Add("SetColumnInTable", SetColumnInTable);
            _elementStartRules.Add("SetPrimaryKeyOfTable", SetPrimaryKeyOfTable);
            _elementStartRules.Add("SetVariable", SetVariable);
            _elementStartRules.Add("StoreExcelDateToVariables", StoreExcelDateToVariables);

            _elementStartRules.Add("CompareVariableWithValue", CompareVariableWithValue);
            _elementEndRules.Add("CompareVariableWithValue", PopConditionalStack);
            _elementStartRules.Add("CompareVariableWithVariable", CompareVariableWithVariable);
            _elementEndRules.Add("CompareVariableWithVariable", PopConditionalStack);
        }

        override public void ExecuteElementStartRule(String actionName, Dictionary<String, String> parameters)
        {
            if (_elementStartRules.ContainsKey(actionName))
            {
                _elementStartRules[actionName](parameters);
            }
        }

        override public void ExecuteElementEndRule(String actionName)
        {
            if (_elementEndRules.ContainsKey(actionName))
            {
                _elementEndRules[actionName]();
            }
        }

        private void AddColumnToTable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Table") && parameters.ContainsKey("Column"))
            {
                if (EngineState.TableList.Exists((x) => { return (x.TableName == parameters["Table"]); }))
                {
                    DataTable table = EngineState.TableList.Find((x) => { return (x.TableName == parameters["Table"]); });
                    if (!table.Columns.Contains(parameters["Column"]))
                    {
                        DataColumn dc = new DataColumn
                        {
                            ColumnName = parameters["Column"]
                        };
                        if (parameters.ContainsKey("MaxLength") && Int32.TryParse(parameters["MaxLength"], out int maxLength))
                        {
                            dc.MaxLength = maxLength;
                        }
                        table.Columns.Add(dc);
                    }
                }
            }
        }

        private void ClearTable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Table"))
            {
                if (EngineState.TableList.Exists((x) => { return (x.TableName == parameters["Table"]); }))
                {
                    EngineState.TableList.Find((x) => { return (x.TableName == parameters["Table"]); }).Clear();
                }
            }
        }

        private void CompareVariableWithValue(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Variable") && parameters.ContainsKey("Value") && EngineState.VariableDictionary.ContainsKey(parameters["Variable"]))
            {
                EngineState.ConditionalStack.Push(parameters["Value"].Equals(EngineState.VariableDictionary[parameters["Variable"]].ToString()));
            }
        }

        private void CompareVariableWithVariable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Variable") && parameters["SecondVariable"] != null && EngineState.VariableDictionary.ContainsKey(parameters["Variable"]) && EngineState.VariableDictionary.ContainsKey(parameters["SecondVariable"]))
            {
                EngineState.ConditionalStack.Push(EngineState.VariableDictionary[parameters["Variable"]].ToString().Equals(EngineState.VariableDictionary[parameters["SecondVariable"]]));
            }
        }

        private void CreateTable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Table"))
            {
                if (!EngineState.TableList.Exists((x) => { return (x.TableName == parameters["Table"]); }))
                {
                    DataTable table = new DataTable
                    {
                        TableName = parameters["Table"]
                    };
                    EngineState.TableList.Add(table);
                }
            }
        }

        private void ExtractVariablesFromTable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Table") && EngineState.TableList.Exists((x) => { return (x.TableName == parameters["Table"]); }) && parameters.ContainsKey("SelectQuery") && parameters.ContainsKey("ParameterVariables") && parameters.ContainsKey("SingleOutputTextFormat") && parameters.ContainsKey("OutputVariable"))
            {
                if (!EngineState.VariableDictionary.ContainsKey(parameters["OutputVariable"]))
                {
                    EngineState.VariableDictionary.Add(parameters["OutputVariable"], String.Empty);
                }
                if (parameters.ContainsKey("CountVariable") && !EngineState.VariableDictionary.ContainsKey(parameters["CountVariable"]))
                {
                    EngineState.VariableDictionary.Add(parameters["CountVariable"], "0");
                }

                DataTable table = EngineState.TableList.Find((x) => { return (x.TableName == parameters["Table"]); });
                List<Object> parameterVariables = new List<Object>();
                foreach (String str in parameters["ParameterVariables"].Split(','))
                {
                    if (EngineState.VariableDictionary.ContainsKey(str))
                    {
                        parameterVariables.Add(EngineState.VariableDictionary[str]);
                    }
                }

                DataRow[] foundRows = table.Select(String.Format(parameters["SelectQuery"], parameterVariables.ToArray<Object>())).Distinct<DataRow>().ToArray<DataRow>();
                StringBuilder strBld = new StringBuilder();
                List<Object> columnsValues = new List<Object>();
                if (parameters.ContainsKey("MultipleOutputTextFormat"))
                {
                    for (int i = 0; i < foundRows.Length; i++)
                    {
                        if (parameters.ContainsKey("Columns"))
                        {
                            columnsValues.Clear();
                            foreach (String str in parameters["Columns"].Split(','))
                            {
                                columnsValues.Add(foundRows[i][foundRows[i].Table.Columns[str]]);
                            }
                        }
                        if (i == (foundRows.Length - 1))
                        {
                            strBld.AppendFormat(parameters["SingleOutputTextFormat"], columnsValues.Count == 0 ? foundRows[i].ItemArray : columnsValues.ToArray<Object>());
                        }
                        else
                        {
                            strBld.AppendFormat(parameters["MultipleOutputTextFormat"], columnsValues.Count == 0 ? foundRows[i].ItemArray : columnsValues.ToArray<Object>());
                        }
                    }
                }
                else
                {
                    if (parameters.ContainsKey("Columns"))
                    {
                        columnsValues.Clear();
                        foreach (String str in parameters["Columns"].Split(','))
                        {
                            columnsValues.Add(foundRows[0][foundRows[0].Table.Columns[str]]);
                        }
                    }
                    strBld.AppendFormat(parameters["SingleOutputTextFormat"], columnsValues.Count == 0 ? foundRows[0].ItemArray : columnsValues.ToArray<Object>());
                }

                EngineState.VariableDictionary[parameters["OutputVariable"]] = strBld.ToString();
                if (parameters.ContainsKey("CountVariable"))
                {
                    EngineState.VariableDictionary[parameters["CountVariable"]] = foundRows.Length;
                }
            }
        }

        private void LookUpTable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Table") && EngineState.TableList.Exists((x) => { return (x.TableName == parameters["Table"]); }) && parameters.ContainsKey("PrimaryKeyVariables") && parameters.ContainsKey("RetrieveColumns") && parameters.ContainsKey("StoreAtVariables"))
            {
                DataTable table = (EngineState.TableList.Find((x) => { return (x.TableName == parameters["Table"]); }));

                String[] primaryKeyVariablesStr = parameters["PrimaryKeyVariables"].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                String[] retrieveColumnsStr = parameters["RetrieveColumns"].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                String[] storeAtVariablesStr = parameters["StoreAtVariables"].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                if (table.PrimaryKey.Length != 0 && (table.PrimaryKey.Length == primaryKeyVariablesStr.Length) && (retrieveColumnsStr.Length == storeAtVariablesStr.Length))
                {
                    bool doAllColumnsExist = true;
                    bool doAllVariablesExist = true;

                    foreach (String str in primaryKeyVariablesStr)
                    {
                        if (!EngineState.VariableDictionary.ContainsKey(str))
                        {
                            doAllVariablesExist = false;
                            break;
                        }
                    }
                    foreach (String str in retrieveColumnsStr)
                    {
                        if (!table.Columns.Contains(str))
                        {
                            doAllColumnsExist = false;
                            break;
                        }
                    }
                    if (doAllColumnsExist && doAllVariablesExist)
                    {
                        foreach (String str in storeAtVariablesStr)
                        {
                            if (!EngineState.VariableDictionary.ContainsKey(str))
                            {
                                EngineState.VariableDictionary.Add(str, null);
                            }
                        }
                        Object[] keys = new Object[table.PrimaryKey.Length];

                        for (int i = 0; i < table.PrimaryKey.Length; i++)
                        {
                            keys[i] = EngineState.VariableDictionary[primaryKeyVariablesStr[i]];
                        }
                        DataRow dr = table.Rows.Find(keys);
                        if (dr != null)
                        {
                            for (int i = 0; i < retrieveColumnsStr.Length; i++)
                            {
                                EngineState.VariableDictionary[storeAtVariablesStr[i]] = dr.ItemArray[table.Columns[retrieveColumnsStr[i]].Ordinal];
                            }
                        }
                    }
                }
            }
        }

        private void LogEvent(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Format"))
            {
                String[] variables = new String[0];

                if (parameters.ContainsKey("Variables"))
                {
                    variables = parameters["Variables"].Split(',');

                    for (int i = 0; i < variables.Length; i++)
                    {
                        if (EngineState.VariableDictionary.ContainsKey(variables[i]))
                        {
                            variables[i] = EngineState.VariableDictionary[variables[i]].ToString();
                        }
                        else
                        {
                            variables[i] = String.Empty;
                        }
                    }
                }
                //using (StreamWriter file = new StreamWriter(verboseLog, true))
                //{
                //    file.WriteLine(String.Format(parameters["Format"], variables));
                //}
            }
        }

        private void MoveVariablesToTable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Table") && EngineState.TableList.Exists((x) => { return (x.TableName == parameters["Table"]); }))
            {
                System.Data.DataTable table = EngineState.TableList.Find((x) => { return (x.TableName == parameters["Table"]); });
                DataRow dataRow = table.NewRow();
                bool willBeInserted = true;

                foreach (KeyValuePair<String, Object> keyPair in EngineState.VariableDictionary)
                {
                    if (dataRow.Table.Columns.Contains(keyPair.Key))
                    {
                        if (keyPair.Value == null || (table.PrimaryKey.Length != 0 && Array.Exists(table.PrimaryKey, (x) => { return (x.ColumnName == keyPair.Key); }) && keyPair.Value.ToString() == String.Empty))
                        {
                            willBeInserted = false;
                            break;
                        }
                        dataRow[keyPair.Key] = keyPair.Value;
                    }
                }
                if (willBeInserted)
                {
                    try
                    {
                        if (table.PrimaryKey.Length != 0)
                        {
                            List<Object> primaryKeys = new List<Object>();
                            foreach (DataColumn dataColumn in table.PrimaryKey)
                            {
                                primaryKeys.Add(dataRow.ItemArray[dataColumn.Ordinal]);
                            }
                            DataRow updatedRow = table.Rows.Find(primaryKeys.ToArray());
                            if (updatedRow == null)
                            {
                                table.Rows.Add(dataRow);
                            }
                            else
                            {
                                for (int i = 0; i < updatedRow.ItemArray.Length; i++)
                                {
                                    if (!updatedRow[i].Equals(dataRow[i]))
                                    {
                                        updatedRow[i] = dataRow[i];
                                    }
                                }
                            }
                        }
                        table.AcceptChanges();
                    }
                    catch (ConstraintException)
                    {
                        // Duplicate record insertion attemps will be ignored
                    }
                }
            }
        }

        private void RemoveAllVariables(Dictionary<String, String> parameters)
        {
            EngineState.VariableDictionary.Clear();
        }

        private void RemoveVariable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Variable") && parameters.ContainsKey("Value"))
            {
                if (EngineState.VariableDictionary.ContainsKey(parameters["Variable"]))
                {
                    EngineState.VariableDictionary.Remove(parameters["Variable"]);
                }
            }
        }

        private void RemoveColumnFromTable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Table") && parameters.ContainsKey("Column"))
            {
                if (EngineState.TableList.Exists((x) => { return (x.TableName == parameters["Table"]); }))
                {
                    DataTable table = EngineState.TableList.Find((x) => { return (x.TableName == parameters["Table"]); });
                    if (table.Columns.Contains(parameters["Column"]))
                    {
                        table.Columns.Remove(parameters["Column"]);
                    }
                }
            }
        }

        private void RemoveTable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Table"))
            {
                if (EngineState.TableList.Exists((x) => { return (x.TableName == parameters["Table"]); }))
                {
                    EngineState.TableList.Remove(EngineState.TableList.Find((x) => { return (x.TableName == parameters["Table"]); }));
                }
            }
        }

        private void SetColumnInTable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Table") && EngineState.TableList.Exists((x) => { return (x.TableName == parameters["Table"]); }) && parameters.ContainsKey("Column") && parameters.ContainsKey("Value") && parameters.ContainsKey("SelectQuery") && parameters.ContainsKey("ParameterVariables"))
            {
                DataTable table = EngineState.TableList.Find((x) => { return (x.TableName == parameters["Table"]); });
                List<Object> criteria = new List<Object>();
                foreach (String str in parameters["ParameterVariables"].Split(','))
                {
                    if (EngineState.VariableDictionary.ContainsKey(str))
                    {
                        criteria.Add(EngineState.VariableDictionary[str]);
                    }
                }
                DataRow[] foundRows = table.Select(String.Format(parameters["SelectQuery"], criteria.ToArray<Object>()));
                for (int i = 0; i < foundRows.Length; i++)
                {
                    foundRows[i][foundRows[i].Table.Columns[parameters["Column"]].Ordinal] = parameters["Value"];
                }
            }
        }

        private void SetPrimaryKeyOfTable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Table") && EngineState.TableList.Exists((x) => { return (x.TableName == parameters["Table"]); }) && parameters.ContainsKey("UniqueConstraintColumns"))
            {
                DataTable table = EngineState.TableList.Find((x) => { return (x.TableName == parameters["Table"]); });
                String[] columnsStr = parameters["UniqueConstraintColumns"].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                List<DataColumn> constraintColumns = new List<DataColumn>();
                bool doAllColumnsExistIntheTable = true;

                foreach (String str in columnsStr)
                {
                    if (table.Columns.Contains(str))
                    {
                        constraintColumns.Add(table.Columns[str]);
                    }
                    else
                    {
                        doAllColumnsExistIntheTable = false;
                        break;
                    }
                }
                if (doAllColumnsExistIntheTable)
                {
                    table.PrimaryKey = constraintColumns.ToArray();
                }
            }
        }

        private void CopyVariable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Variable") && parameters.ContainsKey("IntoVariable") && EngineState.VariableDictionary.ContainsKey(parameters["Variable"]))
            {
                if (EngineState.VariableDictionary.ContainsKey(parameters["IntoVariable"]))
                {
                    EngineState.VariableDictionary[parameters["IntoVariable"]] = EngineState.VariableDictionary[parameters["Variable"]];
                }
                else
                {
                    EngineState.VariableDictionary.Add(parameters["IntoVariable"], EngineState.VariableDictionary[parameters["Variable"]]);
                }
            }
        }

        private void SetVariable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Variable"))
            {
                if (parameters.ContainsKey("Value"))
                {
                    if (EngineState.VariableDictionary.ContainsKey(parameters["Variable"]))
                    {
                        EngineState.VariableDictionary[parameters["Variable"]] = parameters["Value"];
                    }
                    else
                    {
                        EngineState.VariableDictionary.Add(parameters["Variable"], parameters["Value"]);
                    }
                }
                else if (EngineState.VariableDictionary.ContainsKey(parameters["Variable"]))
                {
                    if (parameters.ContainsKey("ConcatenateText"))
                    {
                        EngineState.VariableDictionary[parameters["Variable"]] += parameters["ConcatenateText"];
                    }
                    else if (parameters.ContainsKey("ConcatenateVariable") && EngineState.VariableDictionary.ContainsKey(parameters["ConcatenateVariable"]))
                    {
                        EngineState.VariableDictionary[parameters["Variable"]] += EngineState.VariableDictionary[parameters["ConcatenateVariable"]].ToString();
                    }
                    else if (parameters.ContainsKey("Increment") && Int32.TryParse(parameters["Increment"], out int increment))
                    {
                        EngineState.VariableDictionary[parameters["Variable"]] = Int32.Parse(EngineState.VariableDictionary[parameters["Variable"]].ToString()) + increment;
                    }
                    else if (parameters.ContainsKey("Regex") && parameters.ContainsKey("ReplaceWith"))
                    {
                        EngineState.VariableDictionary[parameters["Variable"]] = Regex.Replace(EngineState.VariableDictionary[parameters["Variable"]].ToString(), parameters["Regex"], parameters["ReplaceWith"]);
                    }
                }
            }
        }

        private void StoreExcelDateToVariables(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Variable") && EngineState.VariableDictionary.ContainsKey(parameters["Variable"]) && parameters.ContainsKey("DayVariable") && parameters.ContainsKey("MonthVariable") && parameters.ContainsKey("YearVariable"))
            {
                if (Int64.TryParse(EngineState.VariableDictionary[parameters["Variable"]].ToString(), out long ticks))
                {
                    DateTime convertedDate = new DateTime(1899, 12, 30).AddDays(ticks); // 29.2.1900 Excel bug ı yüzünden 31 i değil 30 u

                    if (!EngineState.VariableDictionary.ContainsKey(parameters["DayVariable"]))
                    {
                        EngineState.VariableDictionary.Add(parameters["DayVariable"], convertedDate.Day.ToString());
                    }
                    else
                    {
                        EngineState.VariableDictionary[parameters["DayVariable"]] = convertedDate.Day.ToString();
                    }
                    if (!EngineState.VariableDictionary.ContainsKey(parameters["MonthVariable"]))
                    {
                        EngineState.VariableDictionary.Add(parameters["MonthVariable"], convertedDate.Month.ToString());
                    }
                    else
                    {
                        EngineState.VariableDictionary[parameters["MonthVariable"]] = convertedDate.Month.ToString();
                    }
                    if (!EngineState.VariableDictionary.ContainsKey(parameters["YearVariable"]))
                    {
                        EngineState.VariableDictionary.Add(parameters["YearVariable"], convertedDate.Year.ToString());
                    }
                    else
                    {
                        EngineState.VariableDictionary[parameters["YearVariable"]] = convertedDate.Year.ToString();
                    }
                }
            }
        }
    }
}