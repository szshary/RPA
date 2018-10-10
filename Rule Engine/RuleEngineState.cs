using System;
using System.Collections.Generic;
using System.Data;

namespace RPA.Core
{
    public class RuleEngineState
    {
        public List<DataTable> TableCollection;
        public Dictionary<String, Object> VariableCollection;
        public Stack<bool> ConditionalStack;

        public RuleEngineState()
        {
            TableCollection = new List<DataTable>();
            VariableCollection = new Dictionary<String, Object>();
            ConditionalStack = new Stack<bool>();
        }
    }
}