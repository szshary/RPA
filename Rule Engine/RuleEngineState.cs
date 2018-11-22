using System;
using System.Collections.Generic;
using System.Data;

namespace RPA.Core
{
    public class RuleEngineState
    {
        public List<DataTable> TableList;
        public Dictionary<String, Object> VariableDictionary;
        public Stack<bool> ConditionalStack;

        public RuleEngineState()
        {
            TableList = new List<DataTable>();
            VariableDictionary = new Dictionary<String, Object>();
            ConditionalStack = new Stack<bool>();
        }
    }
}