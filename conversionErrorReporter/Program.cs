using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Threading;

namespace conversionErrorReporter
{
    class Program
    {
        static void Main(string[] args)
        {
            ErrorReporter theReporter = new ErrorReporter();
            theReporter.Run();  //the reporter will run at least once if looper = false

            // program will run only once if looperKey != "true"
            string looperPath = ConfigurationManager.AppSettings["looperPath"];
            string looperKey;
            /*NOTE:
             *Var looperPath holds the path to a text document. Var looperKey holds the contents of the text document.
             *The while loop continues indefinitely as long as the file contains the specified string "true". Anything else breaks the loop.
             * String is case insensitive and whitespaces have no effect.
             */
            while (string.Equals(looperKey = File.ReadAllText(looperPath).Trim(), "true", StringComparison.OrdinalIgnoreCase))
            {
                theReporter = new ErrorReporter();
                theReporter.Run();
                Thread.Sleep(5000);
            }
        }
    }
}
