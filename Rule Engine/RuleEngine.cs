using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Xml;
using System.Windows;

namespace RPA.Core
{
    sealed public class RuleEngine : RuleSet
    {
        //private Stopwatch _stopWatch;
        private RuleEngineState _engineState;
        private Stack<RuleSet> _activeRuleSets;

        public RuleEngine(String ruleFilePath) : base()
        {
            Directory.SetCurrentDirectory(ruleFilePath);
            _engineState = new RuleEngineState();
            _activeRuleSets = new Stack<RuleSet>();
            _activeRuleSets.Push(new DefaultRuleSet());

            _elementStartRules.Add("EndlessLoopUntilVariableAndValueEqual", EndlessLoopUntilVariableAndValueEqual);
            _elementStartRules.Add("LoopThroughTable", LoopThroughTable);
            _elementStartRules.Add("ProcessTaskFile", ProcessTaskFile);

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

        override public void ExecuteElementStartRule(String actionName, Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (_elementStartRules.ContainsKey(actionName))
            {
                _elementStartRules[actionName](parameters, engineState);
            }
            _activeRuleSets.Peek().ExecuteElementStartRule(actionName, parameters, engineState);
        }

        override public void ExecuteElementEndRule(String actionName, RuleEngineState engineState)
        {
            _activeRuleSets.Peek().ExecuteElementEndRule(actionName, engineState);
            if (_elementEndRules.ContainsKey(actionName))
            {
                _elementEndRules[actionName](engineState);
            }
        }

        private void AddBrowserRuleSet(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            _activeRuleSets.Push(new BrowserRuleSet(_activeRuleSets.Peek()));
        }

        private void AddDatabaseRuleSet(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            _activeRuleSets.Push(new DatabaseRuleSet(_activeRuleSets.Peek()));
        }

        private void AddExcelRuleSet(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            _activeRuleSets.Push(new ExcelRuleSet(_activeRuleSets.Peek()));
        }

        private void AddTN5250RuleSet(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            _activeRuleSets.Push(new TN5250RuleSet(_activeRuleSets.Peek()));
        }

        private void AddWordRuleSet(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            _activeRuleSets.Push(new WordRuleSet(_activeRuleSets.Peek()));
        }

        private void ShrinkRuleSet(RuleEngineState engineState)
        {
            _activeRuleSets.Pop();
        }

        private void EndlessLoopUntilVariableAndValueEqual(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (parameters.ContainsKey("FileName") && parameters.ContainsKey("Variable") && engineState.VariableCollection.ContainsKey(parameters["Variable"]) && parameters.ContainsKey("Value"))
            {
                while (!engineState.VariableCollection[parameters["Variable"]].ToString().Equals(parameters["Value"]))
                {
                    ProcessTaskFile(parameters, engineState);
                }
            }
        }

        private void LoopThroughTable(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (parameters.ContainsKey("FileName") && parameters.ContainsKey("Table") && engineState.TableCollection.Exists((x) => { return (x.TableName == parameters["Table"]); }))
            {
                foreach (DataRow dr in (engineState.TableCollection.Find((x) => { return (x.TableName == parameters["Table"]); }).Rows))
                {
                    foreach (DataColumn col in dr.Table.Columns)
                    {
                        if (engineState.VariableCollection.ContainsKey(col.ColumnName))
                        {
                            engineState.VariableCollection[col.ColumnName] = dr.ItemArray[col.Ordinal].ToString();
                        }
                    }
                    ProcessTaskFile(parameters, engineState);
                }
            }
        }

        private void ProcessTaskFile(Dictionary<String, String> parameters, RuleEngineState engineState)
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
                                            if (!engineState.ConditionalStack.Peek())
                                            {
                                                readerXML.Skip();
                                            }
                                            break;
                                        case "False":
                                            if (engineState.ConditionalStack.Peek())
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
                                            ExecuteElementStartRule(readerXML.Name, actionParameters, engineState);
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
                                            ExecuteElementEndRule(readerXML.Name, engineState);
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

        public void ProcessTaskFile(String ruleFileName)
        {
            if (File.Exists(Path.Combine(Directory.GetCurrentDirectory(), ruleFileName)))
            {
                Dictionary<String, String> parameters = new Dictionary<String, String>();
                parameters.Add("FileName", ruleFileName);
                try
                {
                    ProcessTaskFile(parameters, _engineState);
                }
                catch (Exception ex)
                {
                }
            }
        }
    }
}