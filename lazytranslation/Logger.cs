using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace lazytranslation
{
    class Logger
    {
        String filePath;

        public Logger(String filePath) {
            this.filePath = filePath;

        }

        public void writeToLog(String loggedString) {
            using (StreamWriter writetext = new StreamWriter(this.filePath))
            {
                writetext.WriteLine(loggedString);
            }
        }
        public void readFromLog() {
            using (StreamReader readtext = new StreamReader(this.filePath))
            {
                string readMeText = readtext.ReadLine();
            }
        }


    }
}
