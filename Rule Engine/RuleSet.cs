using System;
using System.Collections.Generic;

namespace RPA.Core
{
    abstract public class RuleSet
    {
        protected readonly Dictionary<String, Action<Dictionary<String, String>, RuleEngineState>> _elementStartRules;
        protected readonly Dictionary<String, Action<RuleEngineState>> _elementEndRules;

        protected RuleSet()
        {
            _elementStartRules = new Dictionary<String, Action<Dictionary<String, String>, RuleEngineState>>();
            _elementEndRules = new Dictionary<String, Action<RuleEngineState>>();
        }

        abstract public void ExecuteElementStartRule(String actionName, Dictionary<String, String> parameters, RuleEngineState engineState);
        abstract public void ExecuteElementEndRule(String actionName, RuleEngineState engineState);

        protected void PopConditionalStack(RuleEngineState engineState)
        {
            engineState.ConditionalStack.Pop();
        }
    }

    abstract public class RuleSetDecorator : RuleSet
    {
        protected RuleSet _decoratedRuleSet;

        protected RuleSetDecorator(RuleSet decoratedRuleSet) : base()
        {
            _decoratedRuleSet = decoratedRuleSet;
        }

        override public void ExecuteElementStartRule(String actionName, Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (_elementStartRules.ContainsKey(actionName))
            {
                _elementStartRules[actionName](parameters, engineState);
            }
            else
            {
                _decoratedRuleSet.ExecuteElementStartRule(actionName, parameters, engineState);
            }
        }

        override public void ExecuteElementEndRule(String actionName, RuleEngineState engineState)
        {
            if (_elementEndRules.ContainsKey(actionName))
            {
                _elementEndRules[actionName](engineState);
            }
            else
            {
                _decoratedRuleSet.ExecuteElementEndRule(actionName, engineState);
            }
        }
    }
}