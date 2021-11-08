using SolidWorks.Interop.sldworks;
using System;
using System.Diagnostics;

namespace BlueByte.SOLIDWORKS.Extensions
{
    public interface ISOLIDWORKSInstanceManager
    {


        
        SldWorks GetSOLIDWORKSInstanceFromProcessID(int PID);
        SldWorks GetNewInstance(string commandLineParameters = SOLIDWORKSInstanceManager.startSWNoJournalDialogAndSuppressAllDialogs, int timeout = 30);

        /// <summary>
        /// Gets a new instance.
        /// </summary>
        /// <param name="unloadaddinsException">Securiy exception while editing registry keys. Ignored if unloadAddins is set to true.</param>
        /// <param name="unloadAddins">Attempt to unload addins if set to true.</param>
        /// <param name="commandlineArgs">commandline args on how to start SOLIDWORSKS. For a full list of arguments, please see this <see href="https://www.cadoverflow.com/t/solidworks-command-line-arguments/279">thread</see>.</param>
        /// <param name="timeoutInSeconds">The timeout in seconds.</param>
        /// <returns></returns>
        SldWorks GetNewInstance(out Exception unloadaddinsException, bool unloadAddins = false, string commandlineArgs = SOLIDWORKSInstanceManager.startSWNoJournalDialogAndSuppressAllDialogs, int timeoutInSeconds = 30);
        void ReleaseInstance(SldWorks swApp);
        void RestartInstance(ref SldWorks swApp, string commandLineParameters = SOLIDWORKSInstanceManager.startSWNoJournalDialogAndSuppressAllDialogs, int timeout = 30, int attempts = 5);
    }
}