using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using System.Xml;
using System.Windows;

namespace RPA.Core
{
    sealed public class RuleEngine : RuleSet
    {
        //private Stopwatch _stopWatch;
        private Stack<StatefulRuleSet> _activeRuleSets;

        public RuleEngine(String ruleFilePath) : base()
        {
            Directory.SetCurrentDirectory(ruleFilePath);
            _activeRuleSets = new Stack<StatefulRuleSet>();
            _activeRuleSets.Push(new DefaultRuleSet());

            _elementStartRules.Add("LoopUntilEqual", LoopUntilEqual);
            _elementStartRules.Add("LoopThroughTable", LoopThroughTable);
            _elementStartRules.Add("ExecuteTaskFile", ExecuteTaskFile);

            _elementStartRules.Add("BrowserSession", AddBrowserRuleSet);
            _elementEndRules.Add("BrowserSession", ShrinkRuleSet);
            _elementStartRules.Add("DatabaseSession", AddDatabaseRuleSet);
            _elementEndRules.Add("DatabaseSession", ShrinkRuleSet);
            _elementStartRules.Add("ExcelSession", AddExcelRuleSet);
            _elementEndRules.Add("ExcelSession", ShrinkRuleSet);
            _elementStartRules.Add("TN5250Session", AddTN5250RuleSet);
            _elementEndRules.Add("TN5250Session", ShrinkRuleSet);
            _elementStartRules.Add("WordSession", AddWordRuleSet);
            _elementEndRules.Add("WordSession", ShrinkRuleSet);
        }

        override public void ExecuteElementStartRule(String actionName, Dictionary<String, String> parameters)
        {
            if (_elementStartRules.ContainsKey(actionName))
            {
                _elementStartRules[actionName](parameters);
            }
            _activeRuleSets.Peek().ExecuteElementStartRule(actionName, parameters);
        }

        override public void ExecuteElementEndRule(String actionName)
        {
            _activeRuleSets.Peek().ExecuteElementEndRule(actionName);
            if (_elementEndRules.ContainsKey(actionName))
            {
                _elementEndRules[actionName]();
            }
        }

        private void AddBrowserRuleSet(Dictionary<String, String> parameters)
        {
            _activeRuleSets.Push(new BrowserRuleSet(_activeRuleSets.Peek()));
        }

        private void AddDatabaseRuleSet(Dictionary<String, String> parameters)
        {
            _activeRuleSets.Push(new DatabaseRuleSet(_activeRuleSets.Peek()));
        }

        private void AddExcelRuleSet(Dictionary<String, String> parameters)
        {
            _activeRuleSets.Push(new ExcelRuleSet(_activeRuleSets.Peek()));
        }

        private void AddTN5250RuleSet(Dictionary<String, String> parameters)
        {
            _activeRuleSets.Push(new TN5250RuleSet(_activeRuleSets.Peek()));
        }

        private void AddWordRuleSet(Dictionary<String, String> parameters)
        {
            _activeRuleSets.Push(new WordRuleSet(_activeRuleSets.Peek()));
        }

        private void ShrinkRuleSet()
        {
            _activeRuleSets.Pop();
        }

        private void LoopUntilEqual(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("FileName") && parameters.ContainsKey("Variable") && _activeRuleSets.Peek().EngineState.VariableDictionary.ContainsKey(parameters["Variable"]) && (parameters.ContainsKey("Value") ^ (parameters.ContainsKey("OtherVariable") && _activeRuleSets.Peek().EngineState.VariableDictionary.ContainsKey(parameters["OtherVariable"]))))
            {
                if (parameters.ContainsKey("Value"))
                {
                    while (!_activeRuleSets.Peek().EngineState.VariableDictionary[parameters["Variable"]].ToString().Equals(parameters["Value"]))
                    {
                        ExecuteTaskFile(parameters);
                    }
                }
                else
                {
                    while (!_activeRuleSets.Peek().EngineState.VariableDictionary[parameters["Variable"]].ToString().Equals(_activeRuleSets.Peek().EngineState.VariableDictionary[parameters["OtherVariable"]].ToString()))
                    {
                        ExecuteTaskFile(parameters);
                    }
                }
            }
        }

        private void LoopThroughTable(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("FileName") && parameters.ContainsKey("Table") && _activeRuleSets.Peek().EngineState.TableList.Exists((x) => { return (x.TableName == parameters["Table"]); }))
            {
                foreach (DataRow dr in (_activeRuleSets.Peek().EngineState.TableList.Find((x) => { return (x.TableName == parameters["Table"]); }).Rows))
                {
                    foreach (DataColumn col in dr.Table.Columns)
                    {
                        if (_activeRuleSets.Peek().EngineState.VariableDictionary.ContainsKey(col.ColumnName))
                        {
                            _activeRuleSets.Peek().EngineState.VariableDictionary[col.ColumnName] = dr.ItemArray[col.Ordinal].ToString();
                        }
                    }
                    ExecuteTaskFile(parameters);
                }
            }
        }

        private void ExecuteTaskFile(Dictionary<String, String> parameters)
        {
            if (parameters.ContainsKey("FileName") && File.Exists(Path.Combine(Directory.GetCurrentDirectory(), parameters["FileName"])))
            {
                using (XmlTextReader readerXML = new XmlTextReader(Path.Combine(Directory.GetCurrentDirectory(), parameters["FileName"])))
                {
                    try
                    {
                        Dictionary<String, String> actionParameters = new Dictionary<String, String>();
                        while (readerXML.Read())
                        {
                            actionParameters.Clear();
                            for (int attCnt = 0; attCnt < readerXML.AttributeCount; attCnt++)
                            {
                                readerXML.MoveToAttribute(attCnt);
                                actionParameters.Add(readerXML.Name, readerXML.Value);
                            }
                            readerXML.MoveToElement();
                            switch (readerXML.NodeType)
                            {
                                case XmlNodeType.Element:
                                    switch (readerXML.Name)
                                    {
                                        case "True":
                                            if (!_activeRuleSets.Peek().EngineState.ConditionalStack.Peek())
                                            {
                                                readerXML.Skip();
                                            }
                                            break;
                                        case "False":
                                            if (_activeRuleSets.Peek().EngineState.ConditionalStack.Peek())
                                            {
                                                readerXML.Skip();
                                            }
                                            break;
                                        case "Optional":
                                            if (MessageBox.Show(String.Format("Would you like to process {0} optional section?", actionParameters.ContainsKey("Name") ? actionParameters["Name"] : "the following"), "Optional Section Execution Confirmation", MessageBoxButton.YesNo) == MessageBoxResult.No)
                                            {
                                                readerXML.Skip();
                                            }
                                            break;
                                        default:
                                            ExecuteElementStartRule(readerXML.Name, actionParameters);
                                            break;
                                    }
                                    break;
                                case XmlNodeType.EndElement:
                                    switch (readerXML.Name)
                                    {
                                        case "True":
                                        case "False":
                                        case "Optional":
                                            break;
                                        default:
                                            ExecuteElementEndRule(readerXML.Name);
                                            break;
                                    }
                                    break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
        }

        public void ExecuteTaskFile(String taskFileName)
        {
            Dictionary<String, String> parameters = new Dictionary<String, String>();
            parameters.Add("FileName", taskFileName);
            try
            {
                ExecuteTaskFile(parameters);
            }
            catch (Exception ex)
            {
            }
        }

        public void ExecuteConcurrently(String taskFileName, int threadCount)
        {
            Dictionary<String, String> parameters = new Dictionary<String, String>();
            parameters.Add("FileName", taskFileName);
            try
            {
                for (int i = 0; i < threadCount; i++)
                {
                    //Task.Factory.

                }
                ExecuteTaskFile(parameters);
            }
            catch (Exception ex)
            {
            }
        }
    }
}