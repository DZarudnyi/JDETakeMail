using System;
using System.IO;


namespace JDETakeMail
{
    public class Logger
    {
        private string _path;
        private string _filename;
        private string _fullPath;
                

        public string Path
        {
            get => _path;
            set
            {
                string tempPath = value.Trim();
                if (tempPath != String.Empty)
                {
                    tempPath = System.IO.Path.GetDirectoryName(value + @"\");
                }
                else {
                    tempPath = $"{System.IO.Directory.GetCurrentDirectory()}{System.IO.Path.DirectorySeparatorChar}logs";
                }
                
                _path = System.IO.Path.GetDirectoryName(tempPath + @"\");
            }
        }

        public string FileName
        {
            get => _filename;
            set
            {
                string tempFilename = System.IO.Path.GetFileName(value);
                if (tempFilename.Trim() == String.Empty)
                {
                    throw new FormatException("Log filename is empty!");
                }

                if (System.IO.Path.GetExtension(tempFilename) == String.Empty) {
                    tempFilename = $"{tempFilename}.log";
                };
                _filename = tempFilename;
            }
        }

        /// <summary>
        /// Full filename of the current log file
        /// </summary>
        public string FullPath => _fullPath =  $"{Path}{System.IO.Path.DirectorySeparatorChar}{FileName}";


        public Logger(string logFileName = "", string logPath = "")
        {
            Path = logPath;
            FileName = logFileName;
        }

        private void WriteToFile(string s)
        {
            try
            {
                // creating directories if not exist
                System.IO.Directory.CreateDirectory(Path);
                FileInfo f = new FileInfo(FullPath);
                using (StreamWriter swriter = f.AppendText())
                {
                    if (s.Trim() == String.Empty)
                    {
                        swriter.WriteLine("");
                    }
                    else
                    {
                        swriter.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss:ffff} | {s}");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Cannot write log file! Error: " + e.Message);
            }            
        }

        public void WriteLine(string s)
        {
            Console.WriteLine(s);
            WriteToFile(s);
        }

        public void InsertEmptyLine()
        {
            WriteLine(String.Empty);
        }
    }
}
