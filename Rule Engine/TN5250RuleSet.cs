using System;
using System.Collections.Generic;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace RPA.Core
{
    sealed internal class TN5250RuleSet : RuleSetDecorator
    {
        private const int MF_BYCOMMAND = 0x00000000;
        private const int SC_CLOSE = 0xF060;
        private const int SC_MINIMIZE = 0xF020;
        private const int SC_MAXIMIZE = 0xF030;
        private const int SC_SIZE = 0xF000;

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool FreeConsole();

        [DllImport("user32.dll")]
        public static extern int DeleteMenu(IntPtr hMenu, int nPosition, int wFlags);

        [DllImport("user32.dll")]
        private static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);

        [DllImport("kernel32.dll", ExactSpelling = true)]
        private static extern IntPtr GetConsoleWindow();

        private readonly static SortedDictionary<String, byte[]> ConsoleFunctionKeyBytes = new SortedDictionary<String, byte[]>()
        {
            { "F1", new byte[] { 27, 49} }, { "F2", new byte[] { 27, 50} }, { "F3", new byte[] { 27, 51} },
            { "F4", new byte[] { 27, 52} }, { "F5", new byte[] { 27, 53} }, { "F6", new byte[] { 27, 54} },
            { "F7", new byte[] { 27, 55} }, { "F8", new byte[] { 27, 56} }, { "F9", new byte[] { 27, 57} },
            { "F10", new byte[] { 27, 48} }, { "F11", new byte[] { 27, 45} }, { "F12", new byte[] { 27, 61} }
        };

        private readonly static SortedDictionary<byte, char> TurkishCharacters = new SortedDictionary<byte, char>()
        {
            { 33, 'Ğ' }, { 34, 'Ü' }, { 35, 'Ö' }, { 36, 'İ' }, { 64, 'Ş' }, { 92, 'ü' }, { 96, 'ı' },
            { 123, 'ç' },{ 125, 'ğ' }, { 126, 'ö' }, { 162, 'Ç' }, { 164, 'ş' }, { 235, '`' }
        };

        private readonly static SortedDictionary<char, byte> EnglishCharacters = new SortedDictionary<char, byte>()
        {
            { 'Ğ', 33 }, { 'Ü', 34 }, { 'Ö', 35 }, { 'İ', 36 }, { 'Ş', 64 }, { 'ü', 92 }, { 'ı', 96 },
            { 'ç', 123 }, { 'ğ', 125 }, { 'ö', 126 }, { 'Ç', 162 }, { 'ş', 164 }, { '`', 235 }
        };

        private TcpClient _tn5250TelnetClient;
        private NetworkStream _tn5250TelnetStream;
        private byte[] _screenBytes;
        private byte[] _commandBytes;
        private char[,] _consoleScreen;
        private int _numOfBytesRead;

        private struct ConsoleCoordinate
        {
            readonly int _left;
            readonly int _top;

            public ConsoleCoordinate(int left, int top)
            {
                _left = left;
                _top = top;
            }
        }

        private int waitDuration = 0;

        public TN5250RuleSet(RuleSet ruleSet) : base(ruleSet)
        {
            _elementStartRules.Add("StartTN5250Session", StartTN5250Session);
            _elementStartRules.Add("SendFunctionKeyToTN5250", SendFunctionKeyToTN5250);
            _elementStartRules.Add("ScrapeFromTN5250", ScrapeFromTN5250);
            _elementStartRules.Add("WriteToTN5250", WriteToTN5250);

            _elementStartRules.Add("CompareVariableWithTN5250", CompareVariableWithTN5250);
            _elementEndRules.Add("CompareVariableWithTN5250", PopConditionalStack);
            _elementStartRules.Add("CompareValueWithTN5250", CompareValueWithTN5250);
            _elementEndRules.Add("CompareValueWithTN5250", PopConditionalStack);

            _elementStartRules.Add("TN5250Session", StartTN5250Session);
            _elementEndRules.Add("TN5250Session", EndTN5250Session);
        }

        private void CompareValueWithTN5250(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (parameters.ContainsKey("Value") && parameters.ContainsKey("Left") && parameters.ContainsKey("Top"))
            {
                if (parameters.ContainsKey("Length") && Int32.TryParse(parameters["Length"], out int length))
                {
                    engineState.ConditionalStack.Push(parameters["Value"].Equals(ScrapeFromTN5250(Convert.ToInt32(parameters["Left"]), Convert.ToInt32(parameters["Top"]), length)));
                }
                else
                {
                    engineState.ConditionalStack.Push(parameters["Value"].Equals(ScrapeFromTN5250(Convert.ToInt32(parameters["Left"]), Convert.ToInt32(parameters["Top"]), parameters["Value"].Length)));
                }
            }
        }

        private void CompareVariableWithTN5250(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (parameters.ContainsKey("Variable") && engineState.VariableCollection.ContainsKey(parameters["Variable"]) && parameters.ContainsKey("Left") && parameters.ContainsKey("Top") && parameters.ContainsKey("Length"))
            {
                engineState.ConditionalStack.Push(engineState.VariableCollection[parameters["Variable"]].Equals(ScrapeFromTN5250(Convert.ToInt32(parameters["Left"]), Convert.ToInt32(parameters["Top"]), Convert.ToInt32(parameters["Length"]))));
            }
        }

        private void RefreshTN5250()
        {
            Thread.Sleep(200); // İnsanlar görebilsin diye
            bool spaceToUnderscore = false;
            List<ConsoleCoordinate> conCoor = new List<ConsoleCoordinate>();
            _numOfBytesRead = _tn5250TelnetStream.Read(_screenBytes, 0, _screenBytes.Length);

            for (int i = 0; i < 80; i++)
            {
                for (int j = 0; j < 25; j++)
                {
                    _consoleScreen[i, j] = ' ';
                }
            }

            for (int i = 0; i < _numOfBytesRead; i += 1)
            {
                if (i + 1 < _numOfBytesRead && _screenBytes[i] == 27 && _screenBytes[i + 1] == 91)
                {
                    if (i + 3 < _numOfBytesRead && _screenBytes[i + 2] == 55 && _screenBytes[i + 3] == 109)
                    {
                        Console.ForegroundColor = ConsoleColor.Black;
                        Console.BackgroundColor = ConsoleColor.Red;
                        i += 3;
                        continue;
                    }
                    else if (i + 3 < _numOfBytesRead && _screenBytes[i + 2] == 48 && _screenBytes[i + 3] == 109)
                    {
                        // Underscore end
                        Console.ForegroundColor = ConsoleColor.White;
                        Console.BackgroundColor = ConsoleColor.Black;
                        spaceToUnderscore = false;
                        i += 3;
                        continue;
                    }
                    else if (i + 3 < _numOfBytesRead && _screenBytes[i + 2] == 50 && _screenBytes[i + 3] == 74)
                    {
                        Console.Clear();
                        i += 3;
                        continue;
                    }
                    else if (i + 3 < _numOfBytesRead && _screenBytes[i + 2] == 49 && _screenBytes[i + 3] == 109)
                    {
                        Console.ForegroundColor = ConsoleColor.White;
                        i += 3;
                        continue;
                    }
                    else if (i + 3 < _numOfBytesRead && _screenBytes[i + 2] == 52 && _screenBytes[i + 3] == 109)
                    {
                        // Underscore start
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.BackgroundColor = ConsoleColor.Black;
                        spaceToUnderscore = true;
                        i += 3;
                        continue;
                    }
                    else if (i + 5 < _numOfBytesRead && _screenBytes[i + 2] == 52 && _screenBytes[i + 3] == 59 && _screenBytes[i + 4] == 49 && _screenBytes[i + 5] == 109)
                    {
                        // Highlight
                        Console.ForegroundColor = ConsoleColor.Green;
                        i += 5;
                        continue;
                    }
                    else if (i + 5 < _numOfBytesRead && _screenBytes[i + 2] == 52 && _screenBytes[i + 3] == 59 && _screenBytes[i + 4] == 55 && _screenBytes[i + 5] == 109)
                    {
                        // Highlight 
                        Console.BackgroundColor = ConsoleColor.Red;
                        i += 5;
                        continue;
                    }
                    else if (i + 5 < _numOfBytesRead && _screenBytes[i + 2] == 53 && _screenBytes[i + 3] == 59 && _screenBytes[i + 4] == 55 && _screenBytes[i + 5] == 109)
                    {
                        // Highlight
                        Console.BackgroundColor = ConsoleColor.Red;
                        i += 5;
                        continue;
                    }
                    else if (i + 4 < _numOfBytesRead && _screenBytes[i + 2] == 63 && _screenBytes[i + 3] == 51 && _screenBytes[i + 4] == 108)
                    {
                        // Unknown Invisible
                        if (_screenBytes[i + 5] != 27 || _screenBytes[i + 6] != 91 || _screenBytes[i + 7] != 63 || _screenBytes[i + 8] != 55 || _screenBytes[i + 9] != 104)
                        {
                            //using (StreamWriter file = new StreamWriter(verboseLog, true))
                            //{
                            //    file.WriteLine("TN5250 Konsolundan tanımlamayan bir byte dizesi geldi.");
                            //}
                        }
                        i += 4;
                        continue;
                    }
                    else if (i + 4 < _numOfBytesRead && _screenBytes[i + 2] == 63 && _screenBytes[i + 3] == 55 && _screenBytes[i + 4] == 104)
                    {
                        // Unknown Invisible
                        conCoor.Add(new ConsoleCoordinate(Console.CursorLeft, Console.CursorTop));
                        i += 4;
                        continue;
                    }
                    else if (i + 5 < _numOfBytesRead && _screenBytes[i + 3] == 59 && _screenBytes[i + 5] == 72)
                    {
                        Console.SetCursorPosition(Int16.Parse(Convert.ToChar(_screenBytes[i + 4]).ToString()) - 1, Int16.Parse(Convert.ToChar(_screenBytes[i + 2]).ToString()) - 1);
                        i += 5;
                        continue;
                    }
                    else if (i + 6 < _numOfBytesRead && _screenBytes[i + 3] == 59 && _screenBytes[i + 6] == 72)
                    {
                        Console.SetCursorPosition((Int16.Parse(Convert.ToChar(_screenBytes[i + 4]).ToString()) * 10 + Int16.Parse(Convert.ToChar(_screenBytes[i + 5]).ToString())) - 1, Int16.Parse(Convert.ToChar(_screenBytes[i + 2]).ToString()) - 1);
                        i += 6;
                        continue;
                    }
                    else if (i + 6 < _numOfBytesRead && _screenBytes[i + 4] == 59 && _screenBytes[i + 6] == 72)
                    {
                        Console.SetCursorPosition(Int16.Parse(Convert.ToChar(_screenBytes[i + 5]).ToString()) - 1, (Int16.Parse(Convert.ToChar(_screenBytes[i + 2]).ToString()) * 10 + Int16.Parse(Convert.ToChar(_screenBytes[i + 3]).ToString())) - 1);
                        i += 6;
                        continue;
                    }
                    else if (i + 7 < _numOfBytesRead && _screenBytes[i + 4] == 59 && _screenBytes[i + 7] == 72)
                    {
                        Console.SetCursorPosition((Int16.Parse(Convert.ToChar(_screenBytes[i + 5]).ToString()) * 10 + Int16.Parse(Convert.ToChar(_screenBytes[i + 6]).ToString())) - 1, (Int16.Parse(Convert.ToChar(_screenBytes[i + 2]).ToString()) * 10 + Int16.Parse(Convert.ToChar(_screenBytes[i + 3]).ToString())) - 1);
                        i += 7;
                        continue;
                    }
                    else if (i + 5 < _numOfBytesRead && _screenBytes[i + 2] == 49 && _screenBytes[i + 3] == 59 && _screenBytes[i + 4] == 55 && _screenBytes[i + 5] == 109)
                    {
                        Console.ForegroundColor = ConsoleColor.Black;
                        Console.BackgroundColor = ConsoleColor.White;
                        i += 5;
                        continue;
                    }
                }
                else
                {
                    if (spaceToUnderscore && _screenBytes[i] == 32)
                    {
                        Console.Write("_");
                    }
                    else
                    {
                        Console.Write(TranslateCodepage500to1026(_screenBytes[i]));
                    }

                    if (Console.CursorLeft == 0)
                    {
                        if (Console.CursorTop == 0)
                        {
                            _consoleScreen[79, Console.CursorTop] = TranslateCodepage500to1026(_screenBytes[i]);
                        }
                        else
                        {
                            _consoleScreen[79, Console.CursorTop - 1] = TranslateCodepage500to1026(_screenBytes[i]);
                        }
                    }
                    else
                    {
                        _consoleScreen[Console.CursorLeft - 1, Console.CursorTop] = TranslateCodepage500to1026(_screenBytes[i]);
                    }
                }
            }
        }

        private void ScrapeFromTN5250(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (parameters.ContainsKey("Left") && parameters.ContainsKey("Top") && parameters.ContainsKey("Length") && parameters.ContainsKey("Variable"))
            {
                if (!engineState.VariableCollection.ContainsKey(parameters["Variable"]))
                {
                    engineState.VariableCollection.Add(parameters["Variable"], ScrapeFromTN5250(Convert.ToInt32(parameters["Left"]), Convert.ToInt32(parameters["Top"]), Convert.ToInt32(parameters["Length"])));
                }
                else
                {
                    engineState.VariableCollection[parameters["Variable"]] = ScrapeFromTN5250(Convert.ToInt32(parameters["Left"]), Convert.ToInt32(parameters["Top"]), Convert.ToInt32(parameters["Length"]));
                }
            }
        }

        private String ScrapeFromTN5250(int left, int top, int length)
        {
            StringBuilder strBld = new StringBuilder();

            if (left > 80 || top > 25)
            {
                return strBld.ToString();
            }

            left -= 1;
            top -= 1;

            for (int i = left; i < left + length; i++)
            {
                if (i >= 80)
                {
                    i = 0;
                    top += 1;
                    if (top >= 25)
                    {
                        break;
                    }
                }
                strBld.Append(_consoleScreen[i, top]);
            }
            return strBld.ToString().Trim();
        }

        private void SendFunctionKeyToTN5250(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (parameters.ContainsKey("FunctionKey") && ConsoleFunctionKeyBytes.ContainsKey(parameters["FunctionKey"]))
            {
                if (parameters.ContainsKey("WaitDuration"))
                {
                    waitDuration = Convert.ToInt32(parameters["WaitDuration"]);
                }

                _commandBytes = ConsoleFunctionKeyBytes[parameters["FunctionKey"]];
                if (_commandBytes.Length != 0)
                {
                    _tn5250TelnetStream.Write(_commandBytes, 0, _commandBytes.Length);
                    if (waitDuration != 0)
                    {
                        Thread.Sleep(waitDuration);
                    }
                    RefreshTN5250();
                }
            }
        }

        private void StartAS400Console()
        {
            _screenBytes = new byte[10000];
            _consoleScreen = new char[80, 25];

            AllocConsole();

            IntPtr conWin = GetConsoleWindow();
            IntPtr sysMenu = GetSystemMenu(conWin, false);

            if (conWin != IntPtr.Zero)
            {
                DeleteMenu(sysMenu, SC_CLOSE, MF_BYCOMMAND);
                DeleteMenu(sysMenu, SC_MINIMIZE, MF_BYCOMMAND);
                DeleteMenu(sysMenu, SC_MAXIMIZE, MF_BYCOMMAND);
                DeleteMenu(sysMenu, SC_SIZE, MF_BYCOMMAND);
            }

            Console.Title = "Anadolu Hayat Emeklilik - AS400 Emülasyonu";
            Console.SetWindowSize(80, 25);
            Console.SetBufferSize(80, 25);
            Console.SetWindowPosition(0, 0);
            Console.CursorVisible = true;
            Console.InputEncoding = Encoding.UTF8;
            Console.OutputEncoding = Encoding.UTF8;
            Console.ForegroundColor = ConsoleColor.White;
            Console.BackgroundColor = ConsoleColor.Black;

            Console.Clear();
            _numOfBytesRead = _tn5250TelnetStream.Read(_screenBytes, 0, _screenBytes.Length);
            _commandBytes = Encoding.ASCII.GetBytes("\n");
            _tn5250TelnetStream.Write(_commandBytes, 0, _commandBytes.Length);
            Console.SetCursorPosition(0, 0);
            RefreshTN5250();
        }

        private char TranslateCodepage500to1026(byte translatedByte)
        {
            if (TurkishCharacters.ContainsKey(translatedByte))
            {
                return TurkishCharacters[translatedByte];
            }
            else
            {
                return Convert.ToChar(translatedByte);
            }
        }

        private byte TranslateCodepage1026to500(char translatedChar)
        {
            if (EnglishCharacters.ContainsKey(translatedChar))
            {
                return EnglishCharacters[translatedChar];
            }
            else
            {
                return Convert.ToByte(translatedChar);
            }
        }

        private void WriteToTN5250(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            bool isSubmit = false;
            int maxLengthOfField = 0;

            if (parameters.ContainsKey("MaxLength"))
            {
                maxLengthOfField = Convert.ToInt32(parameters["MaxLength"]);
            }
            else if (parameters.ContainsKey("Value"))
            {
                maxLengthOfField = parameters["Value"].Length;
            }
            if (parameters.ContainsKey("Submit"))
            {
                isSubmit = Convert.ToBoolean(parameters["Submit"]);
            }
            if (parameters.ContainsKey("WaitDuration"))
            {
                waitDuration = Convert.ToInt32(parameters["WaitDuration"]);
            }
            if (parameters.ContainsKey("Value"))
            {
                WriteToTN5250(parameters["Value"], isSubmit, maxLengthOfField, waitDuration);
            }
            else if (parameters.ContainsKey("Variable") && engineState.VariableCollection.ContainsKey(parameters["Variable"]))
            {
                WriteToTN5250(engineState.VariableCollection[parameters["Variable"]].ToString(), isSubmit, maxLengthOfField, waitDuration);
            }
        }

        private void WriteToTN5250(string inputCommand, bool isSubmit, int maxLengthOfField, int waitDuration)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.BackgroundColor = ConsoleColor.Black;

            List<byte> inputCommandBytes = new List<byte>();
            foreach (char ch in inputCommand.ToCharArray())
            {
                inputCommandBytes.Add(TranslateCodepage1026to500(ch));
            }

            _commandBytes = inputCommandBytes.ToArray();
            if (_commandBytes.Length != 0)
            {
                _tn5250TelnetStream.Write(_commandBytes, 0, _commandBytes.Length);
                RefreshTN5250();
                _commandBytes = new byte[0];
            }
            if (isSubmit)
            {
                _commandBytes = Encoding.ASCII.GetBytes("\r");
            }
            else if (inputCommand.Length < maxLengthOfField)
            {
                _commandBytes = Encoding.ASCII.GetBytes("\t");
            }
            if (_commandBytes.Length != 0)
            {
                _tn5250TelnetStream.Write(_commandBytes, 0, _commandBytes.Length);
                if (waitDuration != 0)
                {
                    Thread.Sleep(waitDuration);
                }
                RefreshTN5250();
            }
        }

        private void StartTN5250Session(Dictionary<String, String> parameters, RuleEngineState engineState)
        {
            if (_tn5250TelnetClient == null)
            {
                if (parameters.ContainsKey("Host") && parameters.ContainsKey("Port"))
                {
                    _tn5250TelnetClient = new TcpClient(parameters["Host"], Convert.ToInt32(parameters["Port"]));
                    _tn5250TelnetStream = _tn5250TelnetClient.GetStream();
                    StartAS400Console();
                }
            }
        }

        private void EndTN5250Session(RuleEngineState engineState)
        {
            if (_tn5250TelnetClient != null)
            {
                _tn5250TelnetStream.Close();
                _tn5250TelnetClient.Close();
                FreeConsole();
            }
        }
    }
}