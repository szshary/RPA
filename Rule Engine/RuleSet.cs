using System;
using System.Collections.Generic;

namespace RPA.Core
{
    abstract public class RuleSet
    {
        protected readonly Dictionary<String, Action<Dictionary<String, String>>> _elementStartRules;
        protected readonly Dictionary<String, Action> _elementEndRules;

        abstract public void ExecuteElementStartRule(String actionName, Dictionary<String, String> parameters);
        abstract public void ExecuteElementEndRule(String actionName);

        protected RuleSet()
        {   
            _elementStartRules = new Dictionary<String, Action<Dictionary<String, String>>>();
            _elementEndRules = new Dictionary<String, Action>();
        }
    }

    abstract public class StatefulRuleSet : RuleSet
    {
        public readonly RuleEngineState EngineState;
        
        protected StatefulRuleSet() : base()
        {
            EngineState = new RuleEngineState();
        }

        protected StatefulRuleSet(RuleEngineState engineState) : base()
        {
            EngineState = engineState;
        }

        protected void PopConditionalStack()
        {
            EngineState.ConditionalStack.Pop();
        }
    }

    abstract public class StatefulRuleSetDecorator : StatefulRuleSet
    {
        protected StatefulRuleSet _decoratedStatefulRuleSet;

        protected StatefulRuleSetDecorator(StatefulRuleSet decoratedRuleSet) : base(decoratedRuleSet.EngineState)
        {
            _decoratedStatefulRuleSet = decoratedRuleSet;
        }
        override public void ExecuteElementStartRule(String actionName, Dictionary<String, String> parameters)
        {
            if (_elementStartRules.ContainsKey(actionName))
            {
                _elementStartRules[actionName](parameters);
            }
            else
            {
                _decoratedStatefulRuleSet.ExecuteElementStartRule(actionName, parameters);
            }
        }

        override public void ExecuteElementEndRule(String actionName)
        {
            if (_elementEndRules.ContainsKey(actionName))
            {
                _elementEndRules[actionName]();
            }
            else
            {
                _decoratedStatefulRuleSet.ExecuteElementEndRule(actionName);
            }
        }
    }
}