using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.ClearScript;
using Microsoft.ClearScript.Windows;
using System.IO;

namespace LychenBASIC
{
    static class Program
    {
        public static VBScriptEngine vbscriptEngine;
        public static Dictionary<string, object> Settings = new Dictionary<string, object>();

        static int Main(string[] args)
        {
            LoadSettingsDictionary(args);

            if (Settings.ContainsKey("/DEBUG") && Settings["/DEBUG"].GetType() == typeof(bool))
            {
                System.Diagnostics.Debugger.Launch();
            }

            vbscriptEngine = new VBScriptEngine(WindowsScriptEngineFlags.EnableDebugging);

            AddSymbols();

            string script = string.Empty;
            string replFile = string.Empty;

            var hasRepl = Settings.ContainsKey("/REPL");
            if (hasRepl)
            {
                if (Settings["/REPL"].GetType() != typeof(bool))
                {
                    replFile = Settings["/REPL"].ToString();
                }
            }

            if ((int)Settings["$ARGC"] > 0)
            {
                script = Settings["$ARG0"].ToString();
                if (!File.Exists(script))
                {
                    Console.WriteLine($"Script {script} not found.");
                    return 2;
                }
                ConnectoToScriptINI(script);
            }
            else
            {
                if (!hasRepl)
                {
                    Console.WriteLine("No script.");
                    return 1;
                }
                ConnectoToScriptINI("repl.INI"); // FIXME. Put this somewhere useful
            }

            // SetupIncludeFunction();

            if (script != string.Empty)
            {
                try
                {
                    vbscriptEngine.Execute(System.IO.File.ReadAllText(script));
                }
                catch (Exception exception)
                {
                    if (exception is IScriptEngineException scriptException)
                    {
                        Console.WriteLine(exception.Source);
                        Console.WriteLine(scriptException.ErrorDetails);
                    }
                }
                //if (evaluand.GetType() != typeof(Microsoft.ClearScript.VoidResult))
                //{
                //    Console.WriteLine($"{evaluand}");
                //}
            }

            if (hasRepl)
            {
                RunREPL(replFile);
            }

            return 0;
        }

        private static void RunREPL(string fileName)
        {
            string cmd;
            do
            {
                Console.Write(Settings["$PROMPT"]);
                cmd = Console.ReadLine();
                if (cmd == "bye")
                {
                    break;
                }
                if (fileName != string.Empty)
                {
                    File.AppendAllText(fileName, cmd + "\r\n");
                }

                object evaluand;
                try
                {
                    evaluand = vbscriptEngine.ExecuteCommand(cmd);
                }
                catch (ScriptEngineException see)
                {
                    evaluand = "";
                    Console.WriteLine(see.ErrorDetails);
                    Console.WriteLine(see.StackTrace);
                }
                catch (NullReferenceException nre)
                {
                    evaluand = "";
                    Console.WriteLine(nre.Message);
                }
                catch (Exception e)
                {
                    evaluand = "";
                    Console.WriteLine(e.Message);
                }

            } while (cmd != "bye");
        }

        //private static void SetupIncludeFunction()
        //{
        //    var obfusc = "CS" + KeyGenerator.GetUniqueKey(36);
        //
        //    vbscriptEngine.AddHostObject(obfusc, vbscriptEngine);
        //    var includeCode = @"
        //    const console = {
        //        log: function () {
        //          var format;
        //          const args = [].slice.call(arguments);
        //          if (args.length > 1) {
        //            format = args.shift();
        //            args.forEach(function (item) {
        //              format = format.replace('%s', item);
        //            });
        //          } else {
        //            format = args.length === 1 ? args[0] : '';
        //          }
        //          CS.System.Console.WriteLine(format);
        //        }
        //      };
        //    function include(fn) {
        //        if (CSFile.Exists(fn)) {
        //            $SYM$.Evaluate(CSFile.ReadAllText(fn));
        //        } else {
        //            throw fn + ' not found.';
        //        }
        //    }
        //    function include_once(fn) {
        //        if (CSFile.Exists(fn)) {
        //            if (!CSSettings.ContainsKey('included_' + fn)) {
        //                $SYM$.Evaluate(CSFile.ReadAllText(fn));
        //                CSSettings.Add('included_' + fn, true);
        //            } else {
        //                throw fn + ' already included.';
        //            }
        //        } else {
        //            throw fn + ' not found.';
        //        }
        //    }".Replace("$SYM$", obfusc);
        //
        //    vbscriptEngine.Evaluate(includeCode);
        //}

        private static void ConnectoToScriptINI(string script)
        {
            var ini = new INI(Path.ChangeExtension(script, ".INI"));
            vbscriptEngine.AddHostObject("CSScriptINI", ini);
        }

        private static void LoadSettingsDictionary(string[] args)
        {
            var argv = new List<string>();
            var cnt = 0;
            var slashCnt = 0;
            foreach (var arg in args)
            {
                if (arg.StartsWith("/"))
                {
                    slashCnt++;
                    if (arg.Contains(":") || arg.Contains("="))
                    {
                        var lhs = arg.Split(new char[] { ':', '=' }, 2);
                        Settings[lhs[0]] = lhs[1];
                    }
                    else
                    {
                        Settings[arg] = true;
                    }
                }
                else
                {
                    Settings[$"$ARG{cnt}"] = arg;
                    cnt++;
                    argv.Add(arg);
                }
            }
            Settings["$ARGC"] = cnt;
            Settings["$ARGV"] = argv.ToArray<string>();
            Settings["/COUNT"] = slashCnt;
            Settings["$PROMPT"] = "Lychen>";
        }

        private static void Run(string fname)
        {
            if (File.Exists(fname))
            {
                try
                {
                    vbscriptEngine.Execute(File.ReadAllText(fname));
                }
                catch (ScriptEngineException see)
                {
                    Console.WriteLine(see.ErrorDetails);
                }
            }
        }

        private static void AddSymbols()
        {
            AddInternalSymbols(ref vbscriptEngine);
            AddHostSymbols(ref vbscriptEngine);
            AddSystemSymbols(ref vbscriptEngine);
            vbscriptEngine
                .Script
                .Print = (Action<object>)Console.WriteLine;
            // vbscriptEngine
            //     .Script
            //     .Printf = (Action<string, object>)Console.WriteLine;
            vbscriptEngine
                .Script
                .Run = (Action<string>)Run;
        }

        private static void AddInternalSymbols(ref VBScriptEngine vbscriptEngine)
        {
            vbscriptEngine.AddHostObject("VBS", vbscriptEngine);
            vbscriptEngine.AddHostType("CSINI", typeof(INI));
            //vbscriptEngine.AddHostType("CSKeyGenerator", typeof(KeyGenerator));
            vbscriptEngine.AddHostObject("CSSettings", Settings);
            vbscriptEngine.AddHostType("CSLychen", typeof(Program)); // Experimental. No idea if useful or dangerous.
            //vbscriptEngine.AddHostType("CSReflection", typeof(Reflections));
            //vbscriptEngine.AddHostType("CSExtension", typeof(Lychen.Extension));
        }

        private static void AddSystemSymbols(ref VBScriptEngine vbscriptEngine)
        {
            vbscriptEngine.AddHostType("CSFile", typeof(System.IO.File));
            vbscriptEngine.AddHostType("CSConsole", typeof(System.Console));
            vbscriptEngine.AddHostType("CSPath", typeof(System.IO.Path));
            vbscriptEngine.AddHostType("CSDirectory", typeof(System.IO.Directory));
            vbscriptEngine.AddHostType("CSDirectoryInfo", typeof(System.IO.DirectoryInfo));
            vbscriptEngine.AddHostType("CSEnvironment", typeof(System.Environment));
            vbscriptEngine.AddHostType("CSString", typeof(System.String));
            vbscriptEngine.AddHostType("CSDateTime", typeof(System.DateTime));
            vbscriptEngine.AddHostType("CSDebugger", typeof(System.Diagnostics.Debugger));
        }

        private static void AddHostSymbols(ref VBScriptEngine vbscriptEngine)
        {
            vbscriptEngine.AddHostObject("CSExtendedHost", new ExtendedHostFunctions());
            vbscriptEngine.AddHostObject("CSHost", new HostFunctions());
            var htc = new HostTypeCollection();
            foreach (var assembly in new string[] { "mscorlib", "System", "System.Core", "System.Data" /*, "RestSharp",  "WebDriver", "WebDriver.Support" */})
            {
                htc.AddAssembly(assembly);
            }
            if (Settings.ContainsKey("/ASSEMBLIES"))
            {
                var assemblies = Settings["/ASSEMBLIES"].ToString().Split(',');
                foreach (var assembly in assemblies)
                {
                    System.Reflection.Assembly assem;
                    try
                    {
                        assem = System.Reflection.Assembly.LoadFrom(assembly);
                        htc.AddAssembly(assem);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                }
            }
            vbscriptEngine.AddHostObject("CS", htc);
        }
    }
}
