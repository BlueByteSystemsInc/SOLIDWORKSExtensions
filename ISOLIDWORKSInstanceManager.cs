using SolidWorks.Interop.sldworks;
using System;
using System.Diagnostics;

namespace BlueByte.SOLIDWORKS.Extensions
{
    public interface ISOLIDWORKSInstanceManager
    {


        
        SldWorks GetSOLIDWORKSInstanceFromProcessID(int PID);

        /// <summary>
        /// Gets a new instance.
        /// </summary>
        /// <param name="commandlineArgs">commandline args on how to start SOLIDWORSKS. For a full list of arguments, please see this <see href="https://www.cadoverflow.com/t/solidworks-command-line-arguments/279">thread</see>.</param>
        /// <param name="timeoutInSeconds">The timeout in seconds.</param>
        /// <returns></returns>
        SldWorks GetNewInstance(string commandLineParameters = SOLIDWORKSInstanceManager.startSWNoJournalDialogAndSuppressAllDialogs, int timeout = 30);

 
        void ReleaseInstance(SldWorks swApp);
        void RestartInstance(ref SldWorks swApp, string commandLineParameters = SOLIDWORKSInstanceManager.startSWNoJournalDialogAndSuppressAllDialogs, int timeout = 30, int attempts = 5);
    }
}