using BlueByte.SOLIDWORKS.Extensions;
using BlueByte.SOLIDWORKS.Extensions.Helpers;
using BlueByte.SOLIDWORKS.Helpers;
using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Tests
{

    internal class Program
    {
       

        static void Main(string[] args)
        {
            var app = OfficeAppInstanceManager.GetOfficeAppInstanceFromProcessID(16080);
        }

 


    }
    public class OfficeAppInstanceManager
    {

        [DllImport("ole32.dll")]
        public static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable port);

        [DllImport("ole32.dll")]
        public static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);


        /// <summary>
        /// Gets the Office app instance from process identifier.
        /// </summary>
        /// <param name="PID">The process ID.</param>
        /// <returns></returns>
        public static dynamic GetOfficeAppInstanceFromProcessID(int PID)
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

                var formattedGuid = runningObjectName.Replace("{", "").Replace("}", "").Replace("!", "");


                Guid g;
                var ret = Guid.TryParse(formattedGuid, out g);

                if (ret == false)
                    continue;
 
                string progId;
                int result = ProgIDFromCLSID(ref g, out progId);

                Console.WriteLine(progId);

                if (runningObjectName.ToLower().Contains("excel"))
                {
                    object runningObjectVal;
                    runningObjectTable.GetObject(monikers[0], out runningObjectVal);

                    // we should be safe to cast to our "real" solidworks object
                    dynamic swObj = runningObjectVal;

                    if (swObj != null && swObj.GetProcessID() == PID)
                    {
                        return swObj;
                    }
                }
            }


            return null;
        }







        [DllImport("ole32.dll")]
        static extern int ProgIDFromCLSID([In()] ref Guid clsid, [MarshalAs(UnmanagedType.LPWStr)] out string lplpszProgID);
    }
}