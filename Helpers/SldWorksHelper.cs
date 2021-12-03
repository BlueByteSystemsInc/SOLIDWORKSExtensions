﻿using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Permissions;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace BlueByte.SOLIDWORKS.Extensions.Helpers
{


    public enum SHOWWINDOW
    {
        SW_HIDE = 0,
        SW_SHOWNORMAL = 1,
        SW_NORMAL = 1,
        SW_SHOWMINIMIZED = 2,
        SW_SHOWMAXIMIZED = 3,
        SW_MAXIMIZE = 3,
        SW_SHOWNOACTIVATE = 4,
        SW_SHOW = 5,
        SW_MINIMIZE = 6,
        SW_SHOWMINNOACTIVE = 7,
        SW_SHOWNA = 8,
        SW_RESTORE = 9,
        SW_SHOWDEFAULT = 10,
        SW_FORCEMINIMIZE = 11,
        SW_MAX = 11,
    }
    public static class SldWorksHelper
    {

        public const string ADDINS_STARTUP_REG_KEY = @"Software\SolidWorks\AddInsStartup";



        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);


        /// <summary>
        /// Maximizes the SOLIDWORKS window.
        /// </summary>
        /// <returns>True if successful, false if not.</returns>
        /// <param name="swApp">The sw application.</param>
        public static bool MaximizeWindow(this SldWorks swApp)
        {
            var frame = swApp.Frame() as IFrame;
            var handle = new IntPtr(frame.GetHWndx64());
            var ret = ShowWindow(handle, (int)SHOWWINDOW.SW_MAXIMIZE);
            return ret;
        }


        /// <summary>
        /// Performs a zoom to fit.
        /// </summary>
        /// <param name="maximizeWindow">Pointer to the SOLIDWORKS application in order to maximize window before performing zoom to fit.</param>
        /// <param name="modelDoc">The model document.</param>
        /// <remarks>This method makes SOLIDWORKS visible when maximizeWindow is specified.</remarks>
        public static void ViewZoomtofit3(this ModelDoc2 modelDoc, SldWorks maximizeWindow = null)
        {
            if (maximizeWindow != null)
            {
                if (maximizeWindow.Visible == false)
                maximizeWindow.Visible = true; 
                maximizeWindow.MaximizeWindow();
            }
            modelDoc.ViewZoomtofit();

        }

        /// <summary>
        /// Unloads add ins.
        /// </summary>
        /// <param name="disabledAddInGuids">The disabled add in guids.</param>
        /// <returns></returns>
        /// <exception cref="System.Exception">Failed to unload addins due to security permissions</exception>
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

                throw new UnloadAddInsFailedException($"Failed to unload addins due to security permissions", e);
            }
            
            return true;
        }
    }
}
