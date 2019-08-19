using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GMS.LIB.DataAccess
{
    /// <summary>
    /// 
    /// </summary>
    public static class TimeOuts
    {
        private static List<string> _listTimeouts = new List<string>()
        {
            "Timeout Expired".ToUpper()
            , "The semaphore timeout period has expired".ToUpper()
            , "The timeout period elapsed".ToUpper()
            , "ExecuteReader requires an open and available Connection".ToUpper()
            , "There is already an open DataReader associated".ToUpper()
        };

        /// <summary>
        /// 
        /// </summary>
        public static List<string> ListTimeouts
        {
            get { return _listTimeouts; }
        }

        public static bool CheckTimeout(string queryResult, out string timeoutMatched)
        {
            bool result = false;
            timeoutMatched = string.Empty;

            foreach (string pattern in TimeOuts.ListTimeouts)
            {
                if (queryResult.ToUpper().Contains(pattern))
                {
                    result = true;
                    timeoutMatched = pattern;
                }
            }

            return result;
        }

    }
}