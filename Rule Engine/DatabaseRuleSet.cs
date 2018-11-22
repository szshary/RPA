using IBM.Data.DB2.iSeries;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;

namespace RPA.Core
{
    sealed public class DatabaseRuleSet : StatefulRuleSetDecorator
    {
        private iDB2Connection _dB2connection;

        public DatabaseRuleSet(StatefulRuleSet statefulRuleSet) : base(statefulRuleSet)
        {
            _elementStartRules.Add("StoreQueryResultToVariables", StoreQueryResultToVariables);
            _elementStartRules.Add("StoreQueryResultToTable", StoreQueryResultToTable);

            _elementStartRules.Add("CompareValueWithScalarQuery", CompareValueWithScalarQuery);
            _elementEndRules.Add("CompareValueWithScalarQuery", PopConditionalStack);

            _elementStartRules.Add("DatabaseSession", StartDatabaseSession);
            _elementEndRules.Add("DatabaseSession", EndDatabaseSession);
        }

        private void CompareValueWithScalarQuery(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Value") && parameters.ContainsKey("SelectQuery"))
            {
                using (iDB2Command command = new iDB2Command(parameters["SelectQuery"], _dB2connection))
                {
                    if (parameters.ContainsKey("ParameterVariables"))
                    {
                        foreach (String str in parameters["ParameterVariables"].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                        {
                            if (EngineState.VariableCollection.ContainsKey(str))
                            {
                                command.Parameters.AddWithValue(str, EngineState.VariableCollection[str]);
                            }
                        }
                    }
                    try
                    {
                        EngineState.ConditionalStack.Push(parameters["Value"].Equals(command.ExecuteScalar().ToString()));
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
        }

        private void StoreQueryResultToTable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("Table"))
            {
                using (DataTable table = new DataTable(parameters["Table"]))
                {
                    using (iDB2Command command = new iDB2Command(parameters["SelectQuery"], _dB2connection))
                    {
                        command.CommandTimeout = 500000;
                        if (parameters.ContainsKey("ParameterVariables"))
                        {
                            Regex numberRegex = new Regex("^[1-9]{1}[0-9]{0,}([.]{1}[0-9]{1,}){0,1}$");
                            foreach (String str in parameters["ParameterVariables"].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                            {
                                if (EngineState.VariableCollection.ContainsKey(str))
                                {
                                    if (numberRegex.IsMatch(EngineState.VariableCollection[str].ToString()))
                                    {
                                        command.Parameters.AddWithValue(str, EngineState.VariableCollection[str].ToString());
                                    }
                                    else
                                    {
                                        command.Parameters.AddWithValue(str, String.Format("'{0}'", EngineState.VariableCollection[str].ToString()));
                                    }
                                }
                            }
                        }
                        using (iDB2DataReader reader = command.ExecuteReader())
                        {
                            table.Load(reader);
                        }
                        if (EngineState.TableCollection.Exists((x) => { return (x.TableName == table.TableName); }))
                        {
                            EngineState.TableCollection.Remove(EngineState.TableCollection.Find((x) => { return (x.TableName == table.TableName); }));
                        }
                        table.TableName = parameters["Table"];
                        EngineState.TableCollection.Add(table);
                    }
                }
            }
        }

        private void StoreQueryResultToVariables(Dictionary<String, String> parameters)
        {
            using (iDB2Command command = new iDB2Command(parameters["SelectQuery"], _dB2connection))
            {
                command.CommandTimeout = 500000;
                if (parameters.ContainsKey("ParameterVariables"))
                {
                    Regex numberRegex = new Regex("^[1-9]{1}[0-9]{0,}([.]{1}[0-9]{1,}){0,1}$");
                    foreach (String str in parameters["ParameterVariables"].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        if (EngineState.VariableCollection.ContainsKey(str))
                        {
                            if (numberRegex.IsMatch(EngineState.VariableCollection[str].ToString()))
                            {
                                command.Parameters.AddWithValue(str, EngineState.VariableCollection[str].ToString());
                            }
                            else
                            {
                                command.Parameters.AddWithValue(str, String.Format("'{0}'", EngineState.VariableCollection[str].ToString()));
                            }
                        }
                    }
                }
                using (iDB2DataReader reader = command.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            if (EngineState.VariableCollection.ContainsKey(reader.GetName(i)))
                            {
                                EngineState.VariableCollection[reader.GetName(i)] = reader[i];
                            }
                            else
                            {
                                EngineState.VariableCollection.Add(reader.GetName(i), reader[i]);
                            }
                        }
                    }
                }
            }
        }

        private void StartDatabaseSession(Dictionary<String, String> parameters)
        {
            if ((_dB2connection == null || _dB2connection.State == ConnectionState.Closed))
            {
                if (parameters.ContainsKey("Host") && parameters.ContainsKey("User") && parameters.ContainsKey("Password"))
                {
                    StringBuilder strBld = new StringBuilder();
                    strBld.Append("DataSource=");
                    strBld.Append(parameters["Host"]);
                    strBld.Append(";UserID=");
                    strBld.Append(parameters["User"]);
                    strBld.Append(";Password=");
                    strBld.Append(parameters["Password"]);
                    strBld.Append(";Pooling=False;ConnectionTimeout=0;DataCompression=false;DefaultCollection=CHAYDAT;SortSequence=UserSpecified;SortTable=qgpl/trktab;");
                    _dB2connection = new iDB2Connection(strBld.ToString());
                    _dB2connection.Open();
                }
            }
        }

        private void EndDatabaseSession()
        {
            if (_dB2connection != null && _dB2connection.State == ConnectionState.Open)
            {
                _dB2connection.Close();
            }
        }
    }
}