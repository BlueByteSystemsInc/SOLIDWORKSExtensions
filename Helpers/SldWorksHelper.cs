using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace BlueByte.SOLIDWORKS.Extensions.Helpers
{ 
    public static class SldWorksHelper
    {
        /// <summary>
        /// Unload add-ins from the specified SOLIDWORKS application.
        /// </summary>
        /// <param name="swApp"></param>
        /// <param name="waitForAddInsToLoadInSeconds">Use this to wait if SW is starting and add-ins need to finish loading...</param>
        /// <returns>True if all add-ins get unloaded</returns>
        public static bool UnloadAddIns(this SldWorks swApp, int waitForAddInsToLoadInSeconds = 30)
        {
            if (swApp == null)
                throw new ArgumentNullException(nameof(swApp));

            Thread.Sleep(waitForAddInsToLoadInSeconds * 1000);

            var clsIds = new List<string>();
            var registry_key = @"SOFTWARE\SolidWorks\AddInsStartup";

            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(registry_key))
            {
                foreach (string subkey_name in key.GetSubKeyNames())
                {
                    if (swApp.GetAddInObject(subkey_name) != null)
                    {
                        clsIds.Add(subkey_name);

                    }
                }
            }


            var files = new List<string>();

            foreach (var clsid in clsIds)
            {


                using (RegistryKey key = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry64).OpenSubKey($@"CLSID\{clsid}"))
                {
                    foreach (string subkey_name in key.GetSubKeyNames())
                    {
                        if (subkey_name == "InprocServer32")
                        {
                            using (RegistryKey subkey = key.OpenSubKey(subkey_name))
                            {
                                var names = new string[] { "", "CodeBase" };
                                foreach (var name in names)
                                {
                                    string path = subkey.GetValue(name) as string;

                                    if (string.IsNullOrWhiteSpace(path))
                                        continue;

                                    var dir = System.IO.Path.GetDirectoryName(path);
                                    var fileName = System.IO.Path.GetFileName(path);

                                    if (string.IsNullOrWhiteSpace(dir))
                                        continue;

                                    if (string.IsNullOrWhiteSpace(fileName))
                                        continue;

                                    if (dir.ToLower().StartsWith("file:\\"))
                                        dir = dir.Replace("file:\\", "");

                                    try
                                    {
                                        var f = new FileInfo(System.IO.Path.Combine(dir, fileName));

                                        if (f.Exists)
                                        {
                                            if (files.Exists(x => x == f.FullName) == false)
                                            {

                                                files.Add(f.FullName);

                                            }
                                        }
                                    }
                                    catch (Exception)
                                    {


                                    }

                                }

                            }

                        }
                    }
                }
            }



            foreach (var file in files)
            {
                var retInt = swApp.UnloadAddIn(file);
                if (retInt == -1)
                    throw new Exception($"Failed to unload {file}");
            }


            return true;
        }
    }
}
