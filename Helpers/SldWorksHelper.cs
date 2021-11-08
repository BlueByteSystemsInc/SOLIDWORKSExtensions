using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Security.Permissions;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace BlueByte.SOLIDWORKS.Extensions.Helpers
{ 
    public static class SldWorksHelper
    {

        public const string ADDINS_STARTUP_REG_KEY = @"Software\SolidWorks\AddInsStartup";

        public static bool UnloadAddIns(out List<string> disabledAddInGuids)
        {
            try
            {
                const int DISABLE_VAL = 0;
                const int ENABLE_VAL = 1;

                disabledAddInGuids = new List<string>();

                var addinsStartup = Registry.CurrentUser.OpenSubKey(ADDINS_STARTUP_REG_KEY, true);

                if (addinsStartup != null)
                {
                    var addInKeyNames = addinsStartup.GetSubKeyNames();

                    if (addInKeyNames != null)
                    {
                        foreach (var addInKeyName in addInKeyNames)
                        {
                            var addInKey = addinsStartup.OpenSubKey(addInKeyName, true);

                            int enableVal;

                            if (int.TryParse(addInKey.GetValue("")?.ToString(), out enableVal))
                            {
                                var loadOnStartup = enableVal == ENABLE_VAL;

                                if (loadOnStartup)
                                {
                                    addInKey.SetValue("", DISABLE_VAL);
                                    disabledAddInGuids.Add(addInKeyName);
                                }
                            }
                        }
                    }
                }

            }
            catch (SecurityException e)
            {

                throw new Exception($"Failed to unload addins due to security permissions", e);
            }
            
            return true;
        }
    }
}
