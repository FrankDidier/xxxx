using ArtLogics.TestSuite.Actions.Conditions;
using ArtLogics.TestSuite.Environment.GlobalVariables;
using ArtLogics.TestSuite.Testing.Actions.GlobalVariables;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArtLogics.Translation.Parser.Utils
{
    public static class ProjectUtilGlobalVariableFunction
    {
        public static void AddGlobalVariableActionItem(NamedGlobalVariable variable1, NamedGlobalVariable variable2,
            OperationKind operation, GlobalVariableAction globalVariableAction)
        {
            var globalVariableActionItem = new GlobalVariableActionItem();

            globalVariableActionItem.VariableA = variable1;
            globalVariableActionItem.VariableB = variable2;
            globalVariableActionItem.Operation = operation;
            globalVariableAction.Items.Add(globalVariableActionItem);
        }

        public static void AddGlobalVariableActionItem(NamedGlobalVariable variable1, decimal value,
            OperationKind operation, GlobalVariableAction globalVariableAction)
        {
            var globalVariableActionItem = new GlobalVariableActionItem();

            globalVariableActionItem.VariableA = variable1;
            globalVariableActionItem.ValueDecimal = value;
            globalVariableActionItem.Operation = operation;
            globalVariableActionItem.State = GlobalVariableItemState.Decimal;
            globalVariableAction.Items.Add(globalVariableActionItem);
        }

        public static void AddGlobalVariableActionItem(NamedGlobalVariable variable1, string value,
    OperationKind operation, GlobalVariableAction globalVariableAction)
        {
            var globalVariableActionItem = new GlobalVariableActionItem();

            globalVariableActionItem.VariableA = variable1;
            globalVariableActionItem.ValueText = value;
            globalVariableActionItem.Operation = operation;
            globalVariableActionItem.State = GlobalVariableItemState.Hex;
            globalVariableAction.Items.Add(globalVariableActionItem);
        }

        public static void AddGlobalVariableComparisonItem (NamedGlobalVariable variable1, NamedGlobalVariable variable2,
            ComparisonOperator operation, GVCompareAction globalVariableComparison)
        {
            var globalVariableConditionItem = new GlobalVariableConditionItem();
            globalVariableConditionItem.Operator = operation;
            globalVariableConditionItem.VariableA = variable1;
            globalVariableConditionItem.VariableB = variable2;
            globalVariableConditionItem.State = GlobalVariableItemState.VarB;
            globalVariableComparison.Condition = globalVariableConditionItem;
        }
    }
}
