using System;
using System.Collections.Generic;
using System.Data;
using System.Xml;

namespace RPA.Core
{
    public class RuleEngineState
    {
        public readonly List<DataTable> TableList;
        public readonly Dictionary<String, Object> VariableDictionary;
        public readonly Stack<bool> ConditionalStack;
        public readonly XmlDocument TaskXmlDocument;

        public RuleEngineState()
        {
            TableList = new List<DataTable>();
            VariableDictionary = new Dictionary<String, Object>();
            ConditionalStack = new Stack<bool>();
            TaskXmlDocument = new XmlDocument();
        }
    }
}