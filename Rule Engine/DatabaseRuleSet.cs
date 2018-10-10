using IBM.Data.DB2.iSeries;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace RPA.Core
{
    sealed public class DatabaseRuleSet : RuleSetDecorator
    {
        private iDB2Connection _dB2connection;

        public DatabaseRuleSet(RuleSet _ruleSet) : base(_ruleSet)
        {
            _elementStartRules.Add("StoreQueryResultToVariables", StoreQueryResultToVariables);
            _elementStartRules.Add("StoreQueryResultToTable", StoreQueryResultToTable);

            _elementStartRules.Add("CompareValueWithScalarQuery", CompareValueWithScalarQuery);
            _elementEndRules.Add("CompareValueWithScalarQuery", PopConditionalStack);

            _elementStartRules.Add("DatabaseSession", StartDatabaseSession);
            _elementEndRules.Add("DatabaseSession", EndDatabaseSession);
        }

        private void CompareValueWithScalarQuery(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (parameters.ContainsKey("Value") && parameters.ContainsKey("SelectQuery"))
            {
                using (iDB2Command command = new iDB2Command(parameters["SelectQuery"], _dB2connection))
                {
                    if (parameters.ContainsKey("ParameterVariables"))
                    {
                        foreach (String str in parameters["ParameterVariables"].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                        {
                            if (engineState.VariableCollection.ContainsKey(str))
                            {
                                command.Parameters.AddWithValue(str, engineState.VariableCollection[str]);
                            }
                        }
                    }
                    try
                    {
                        engineState.ConditionalStack.Push(parameters["Value"].Equals(command.ExecuteScalar().ToString()));
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
        }

        private void StoreQueryResultToTable(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (parameters.ContainsKey("Table"))
            {
                using (DataTable table = new DataTable(parameters["Table"]))
                {
                    using (iDB2DataAdapter adapter = new iDB2DataAdapter(parameters["SelectQuery"], _dB2connection))
                    {
                        adapter.SelectCommand.CommandTimeout = 500000;
                        if (parameters.ContainsKey("ParameterVariables"))
                        {
                            foreach (String str in parameters["ParameterVariables"].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                            {
                                if (engineState.VariableCollection.ContainsKey(str))
                                {
                                    adapter.SelectCommand.Parameters.Add(str, engineState.VariableCollection[str]);
                                }
                            }
                        }
                        adapter.Fill(table);
                        if (engineState.TableCollection.Exists((x) => { return (x.TableName == table.TableName); }))
                        {
                            engineState.TableCollection.Remove(engineState.TableCollection.Find((x) => { return (x.TableName == table.TableName); }));
                        }
                        table.TableName = parameters["Table"];
                        engineState.TableCollection.Add(table);
                    }
                }
            }
        }

        private void StoreQueryResultToVariables(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            using (iDB2Command command = new iDB2Command(parameters["SelectQuery"], _dB2connection))
            {
                command.CommandTimeout = 500000;
                if (parameters.ContainsKey("ParameterVariables"))
                {
                    foreach (String str in parameters["ParameterVariables"].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        if (engineState.VariableCollection.ContainsKey(str))
                        {
                            command.Parameters.AddWithValue(str, engineState.VariableCollection[str].ToString());
                        }
                    }
                }
                using (iDB2DataReader reader = command.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            if (engineState.VariableCollection.ContainsKey(reader.GetName(i)))
                            {
                                engineState.VariableCollection[reader.GetName(i)] = reader[i];
                            }
                            else
                            {
                                engineState.VariableCollection.Add(reader.GetName(i), reader[i]);
                            }
                        }
                    }
                }
            }
        }

        private void StartDatabaseSession(Dictionary<String, String> parameters, RuleEngineState engineState)
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

        private void EndDatabaseSession(RuleEngineState engineState)
        {
            if (_dB2connection != null && _dB2connection.State == ConnectionState.Open)
            {
                _dB2connection.Close();
            }
        }
    }
}