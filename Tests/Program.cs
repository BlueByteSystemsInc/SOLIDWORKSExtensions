using BlueByte.SOLIDWORKS.Extensions;
using BlueByte.SOLIDWORKS.Helpers;
using SolidWorks.Interop.sldworks;
using System;
using System.Diagnostics;
using System.Linq;

namespace Tests
{
    internal class Program
    {
        #region Private Methods

        private static void Main(string[] args)
        {

            var amenLocalFileName = @"C:\SOLIDWORKSPDM\Bluebyte\API\Knapheide\a_assemblies\A_80017120.sldasm";
            var fileName = @"C:\Users\alexe\Desktop\TestAssemlby.SLDASM";
            var manager = new SOLIDWORKSInstanceManager();
            var swApp = default(SldWorks);

            var openRet = default(Tuple<bool, ModelDoc2>);
            var processes = System.Diagnostics.Process.GetProcessesByName("sldworks");

            if (processes != null)
            {
#if Local_Amen
                fileName = amenLocalFileName;
#endif 

                var process = processes.FirstOrDefault();
                if (process != null)
                {
                    swApp = manager.GetSOLIDWORKSInstanceFromProcessID(process.Id);
                    swApp.CommandInProgress = true;
                    if (swApp != null)
                    {
                        swApp.Visible = true;
                    }
                    else return;
                }
                else
                {
                    swApp = manager.GetNewInstance("", 60);
                    swApp.CommandInProgress = true;
                    swApp.Visible = true;
                }
            }

            string[] errors;
            string[] warnings;
            var swModel = swApp.ActiveDoc as ModelDoc2;
            if (swModel == null)
            {
                openRet = swApp.OpenDocument(fileName, out errors, out warnings, SolidWorks.Interop.swconst.swOpenDocOptions_e.swOpenDocOptions_Silent);
                Tests.Do(swApp, openRet.Item2);
                swApp.CommandInProgress = true;
            }
            else
            {
                Tests.Do(swApp, swModel);
                swApp.CommandInProgress = false;
            }

        }

#endregion
    }

    public class Tests
    {
#region Public Methods

        public static void Do(SldWorks swApp)
        {
        }

        public static void Do(SldWorks swApp, ModelDoc2 modelDoc2)
        {

            var timer = new Stopwatch();

            timer.Start();

            swApp.CommandInProgress = true;

            var component = ((modelDoc2 as AssemblyDoc).GetComponents(true) as Object[]).Cast<Component2>().FirstOrDefault();
            modelDoc2.CreateRefPlanesAroundBoundingBoxAroundComponent(component);



            swApp.CommandInProgress = false;

            timer.Stop();


            Console.WriteLine(timer.Elapsed.ToString());

        }

#endregion
    }
}