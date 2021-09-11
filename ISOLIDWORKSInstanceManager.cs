using SolidWorks.Interop.sldworks;
using System.Diagnostics;

namespace BlueByte.SOLIDWORKS.Extensions
{
    public interface ISOLIDWORKSInstanceManager
    {


        
        SldWorks GetSOLIDWORKSInstanceFromProcessID(int PID);
        SldWorks GetNewInstance(string commandLineParameters ="", int timeout = 30);
        void ReleaseInstance(SldWorks swApp);
        void RestartInstance(ref SldWorks swApp, string commandLineParameters = "", int timeout = 30, int attempts = 5);
    }
}