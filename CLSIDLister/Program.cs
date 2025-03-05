

/// CLSIDLister
/// By GitHub@win-lukezhang
/// This is an open-source project, using GNU GPL v3.0 license.
/// Open-Source Repo:
/// https://github.com/win-lukezhang/CLSIDLister


///    GNU GPL v3.0 License

///    CLSIDLister - A tool to list all COM components on Windows.
///    Copyright (C) 2025 win-lukezhang

///    This program is free software: you can redistribute it and/or modify
///    it under the terms of the GNU General Public License as published by
///    the Free Software Foundation, either version 3 of the License, or
///    (at your option) any later version.

///    This program is distributed in the hope that it will be useful,
///    but WITHOUT ANY WARRANTY; without even the implied warranty of
///    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
///    GNU General Public License for more details.

///    You should have received a copy of the GNU General Public License
///    along with this program.  If not, see <https://www.gnu.org/licenses/>.


using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Win32;

namespace CLSIDLister
{
    class Program
    {

        public static readonly string Version = "1.0.0";
        public static readonly string Year = "2025";

        static void Main(string[] args)
        {
            string outputFilePath = string.Empty;
            bool showHelp = false;

            foreach (var arg in args)
            {
                if (arg == "/?" || arg == "--help" || arg == "-?")
                {
                    showHelp = true;
                    break;
                }
            }

            if (showHelp)
            {
                ShowHelp();
                return;
            }

            for (int i = 0; i < args.Length; i++)
            {
                if ((args[i] == "--output" || args[i] == "-o") && i + 1 < args.Length)
                {
                    outputFilePath = args[i + 1];
                }
            }

            List<COMComponent> comComponents = GetCOMComponents();

            if (!string.IsNullOrEmpty(outputFilePath))
            {
                SaveToFile(comComponents, outputFilePath);
                Console.WriteLine($"结果已保存到文件: {outputFilePath}");
            }
            else
            {
                foreach (var component in comComponents)
                {
                    Console.WriteLine($"CLSID: {component.CLSID}");
                    Console.WriteLine($"用途: {component.Description}");
                    Console.WriteLine();
                }
            }
        }

        public static List<COMComponent> GetCOMComponents()
        {
            List<COMComponent> components = new List<COMComponent>();
            string registryKey = @"SOFTWARE\Classes\CLSID";

            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(registryKey))
            {
                if (key != null)
                {
                    foreach (string subkeyName in key.GetSubKeyNames())
                    {
                        using (RegistryKey subkey = key.OpenSubKey(subkeyName))
                        {
                            if (subkey != null)
                            {
                                string description = (string)subkey.GetValue(null);
                                string inprocServer = (string)subkey.OpenSubKey("InprocServer32")?.GetValue(null);

                                components.Add(new COMComponent
                                {
                                    CLSID = subkeyName,
                                    Description = description ?? "无描述",
                                    InprocServer = inprocServer ?? "无 InprocServer32"
                                });
                            }
                        }
                    }
                }
            }

            return components;
        }

        public static void SaveToFile(List<COMComponent> components, string filePath)
        {
            using (StreamWriter writer = new StreamWriter(filePath))
            {
                foreach (var component in components)
                {
                    writer.WriteLine($"CLSID: {component.CLSID}");
                    writer.WriteLine($"用途: {component.Description}");
                    writer.WriteLine();
                }
            }
        }

        public static void ShowHelp()
        {
            Console.WriteLine($"CLSIDLister {Version}");
            Console.WriteLine($"© {Year} win-lukezhang");
            Console.WriteLine("Using GNU GPL v3 License");
            Console.WriteLine("GitHub Repo: https://github.com/win-lukezhang/CLSIDLister");
            Console.WriteLine();
            Console.WriteLine("用法: CLSIDLister [选项] | CLSIDLister");
            Console.WriteLine("选项:");
            Console.WriteLine("  -o, --output <文件路径>   将结果保存到指定文件");
            Console.WriteLine("  /?, --help, -?            显示此帮助信息");
            Console.WriteLine("示例:");
            Console.WriteLine("  CLSIDLister");
            Console.WriteLine("  CLSIDLister -o components.txt");
            Console.WriteLine("  CLSIDLister --output components.txt");
        }

        public class COMComponent
        {
            public string CLSID { get; set; }
            public string Description { get; set; }
            public string InprocServer { get; set; }
        }
    }
}
