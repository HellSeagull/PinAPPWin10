using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace PinAPP
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                RegistryKey getCommandHandlerValue = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\Windows.taskbarpin", RegistryKeyPermissionCheck.ReadSubTree);
                string CommandHandler = getCommandHandlerValue.GetValue("ExplorerCommandHandler").ToString();
                getCommandHandlerValue.Close();

                RegistryKey fixPath = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Classes", RegistryKeyPermissionCheck.ReadWriteSubTree);
                fixPath = fixPath.CreateSubKey("*", RegistryKeyPermissionCheck.ReadWriteSubTree);
                fixPath = fixPath.CreateSubKey("shell", RegistryKeyPermissionCheck.ReadWriteSubTree);
                fixPath = fixPath.CreateSubKey("PinToTaskBar", RegistryKeyPermissionCheck.ReadWriteSubTree);
                fixPath.SetValue("ExplorerCommandHandler", CommandHandler, RegistryValueKind.String);
                fixPath.Close();

                foreach (string fileName in args)
                {
                    if (!File.Exists(fileName))
                    {
                        Console.WriteLine("File " + fileName + " not found");
                        Console.ReadLine();
                        return;
                    }

                    dynamic shellApplication = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
                    dynamic directory = shellApplication.NameSpace(Path.GetDirectoryName(fileName));
                    dynamic link = directory.ParseName(Path.GetFileName(fileName)).InvokeVerb("PinToTaskBar");

                    Console.WriteLine(Path.GetFileName(fileName) + " Pinned to TaskBar !");
                }

                fixPath = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Classes", RegistryKeyPermissionCheck.ReadWriteSubTree);
                fixPath = fixPath.CreateSubKey("*", RegistryKeyPermissionCheck.ReadWriteSubTree);
                fixPath = fixPath.CreateSubKey("shell", RegistryKeyPermissionCheck.ReadWriteSubTree);
                fixPath.DeleteSubKey("PinToTaskBar", true);
                fixPath.Close();
            }
            catch(Exception exp) { Console.WriteLine(exp.ToString()); Console.ReadLine();  return; }

            Console.WriteLine("Press Any Key to exit.");
            Console.ReadLine();
        }
    }
}
