using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CreatingDWG
{
    class LogFile
    {
        private static  List<string> _logList = new List<string>();

        public LogFile()
        {

        }

        public static List<string> LogList
        {
            get { return _logList; }
        }

        public static void addLine(string logLine)
        {
            _logList.Add(logLine);
        }


    }

}
