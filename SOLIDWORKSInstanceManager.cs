﻿using BlueByte.SOLIDWORKS.Extensions.Helpers;
using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security;
using System.Security.Principal;
using System.Threading;

namespace BlueByte.SOLIDWORKS.Extensions
{


    public partial class SOLIDWORKSInstanceManager : ISOLIDWORKSInstanceManager
    {
        public const string startSWNoJournalDialogAndSuppressAllDialogs = "/r";

        public const string startSWBatch = "/b";

        [DllImport("ole32.dll")]
        public static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable port);

        [DllImport("ole32.dll")]
        public static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);


        /// <summary>
        /// Gets the SOLIDWORKS instance from process identifier.
        /// </summary>
        /// <param name="PID">The process ID.</param>
        /// <returns></returns>
        public SldWorks GetSOLIDWORKSInstanceFromProcessID(int PID)
        {
            IntPtr numFetched = IntPtr.Zero;
            IRunningObjectTable runningObjectTable;
            IEnumMoniker monikerEnumerator;
            IMoniker[] monikers = new IMoniker[1];

            GetRunningObjectTable(0, out runningObjectTable);
            runningObjectTable.EnumRunning(out monikerEnumerator);
            IBindCtx b;

            monikerEnumerator.Reset();

            while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
            {
                IBindCtx ctx;
                CreateBindCtx(0, out ctx);

                string runningObjectName;
            
                monikers[0].GetDisplayName(ctx, null, out runningObjectName);


               
                if (runningObjectName.ToLower().Contains("solidworks"))
                {
                    object runningObjectVal;
                    runningObjectTable.GetObject(monikers[0], out runningObjectVal);

                    // we should be safe to cast to our "real" solidworks object
                    SldWorks swObj = runningObjectVal as SldWorks;

                    if (swObj != null && swObj.GetProcessID() == PID)
                    {
                        return swObj;
                    }
                }
            }
            
           
            return null;
        }


        

        /// <summary>
        ///  Creates a new instance of SOLIDWORKS
        /// </summary>
        /// <param name="suppressDialog">Suppress any dialogs.</param>
        /// <param name="timeout">30 seconds for time out.</param>
        /// <returns></returns>
        public SldWorks GetNewInstance(string commandlineParameters = startSWNoJournalDialogAndSuppressAllDialogs, Year_e year = Year_e.Latest, int timeout = 30)
        {
            var swApp = Extension.CreateSldWorks(commandlineParameters, year, timeout);
            return swApp;
        }
     

      

        /// <summary>
        /// Attempts to restart SOLIDWORKS.
        /// </summary>
        /// <param name="swApp"></param>
        public void RestartInstance(ref SldWorks swApp, string commandLineParameters = startSWNoJournalDialogAndSuppressAllDialogs, Year_e year_ = Year_e.Latest, int timeout = 30, int attempts = 5)
        {
            if (attempts <= 2)
                attempts = 5;
            if (swApp != null)
            {
                swApp.CloseAllDocuments(true);
                swApp.ExitApp();
                ReleaseInstance(swApp);
                swApp = null;

            }
            for (int i = 1; i <= attempts; i++)
            {
                try
                {
                    swApp =  GetNewInstance(commandLineParameters, year_, timeout);
                    if (swApp != null)
                        break;
                }
                catch (System.TimeoutException e)
                {

                }
            }
        }


        

        /// <summary>
        /// Releases sldworks from memory since it is a com object not managed by garbage collector.
        /// </summary>
        /// <param name="swApp"></param>
        public void ReleaseInstance(SldWorks swApp)
        {
            if (swApp != null)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(swApp);
        }
    }
}
