using System;
using Microsoft.Win32;

namespace UnregisterComObject {
    static class Program {

        static void Main(string[] args) {
            if (args.Length != 1) {
                Console.WriteLine("Please provide the name of the COM object to remove.");
                return;
            }
            var comObjName = args[0];

            var regKey = Registry.ClassesRoot.OpenSubKey(comObjName + @"\CLSID", false);
            if (regKey == null) {
                Console.WriteLine("COM object not found.");
                return;
            }

            var clsId = regKey.GetValue(string.Empty);
            Registry.ClassesRoot.DeleteSubKeyTree(comObjName);
            Registry.ClassesRoot.DeleteSubKeyTree(@"CLSID\" + clsId);
            Console.WriteLine("COM object uninstalled successfully.");
        }
    }
}
