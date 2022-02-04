using System;
using System.Collections.Generic;
using System.Diagnostics;
using Zirpl.CalcEngine;

namespace SorterToolLibrary.Output_Files.Sorter_Tool_Outputs
{
    public static class SimpleExpressionEvaluator
    {
        #region Methods

        public static CalculationEngine ce = new CalculationEngine();

        /// <summary>
        /// Add-in function for CalcEngine
        /// </summary>
        /// <param name="p">A list of parameters</param>
        /// <returns></returns>
        private static object Contains(List<Expression> p)
        {
            // TODO - remove all the duplicate 'Contains' functions used around the code.
            return ((string)p[0]).Contains((string)p[1]);
        }

        public static bool EvaluateUsingCalcEngine(string expression)
        {
            if (expression == string.Empty)
            {
                return true;
            }
            else
            {
                try
                {
                    var value = (bool)ce.Evaluate(expression);
                    return value;
                }
                catch (Exception)
                {
                    Trace.TraceError("Error parsing Rule '{expression}'. Rule cannot be evaluated.");
                    return false;
                }
            }
        }

        static SimpleExpressionEvaluator()
        {
            ce.RegisterFunction("Contains", 2, Contains);
        }

        #endregion Methods
    }
}