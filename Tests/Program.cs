using BlueByte.SOLIDWORKS.Extensions;
using BlueByte.SOLIDWORKS.Extensions.Helpers;
using BlueByte.SOLIDWORKS.Helpers;
using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
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
                    
                    swApp = manager.GetNewInstance("/b /r", SOLIDWORKSInstanceManager.Year_e.Year2022);
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
                Tests.Do(swApp, swApp.ActiveDoc as ModelDoc2);
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


            var hiddenComponents = modelDoc2.GetHiddenComponents();

            var hiddenComponentObjects = modelDoc2.GetHiddenComponentObjects();
            Console.WriteLine($"Selecting {hiddenComponents.Length} components...");

          modelDoc2.ClearSelection();
            timer = new Stopwatch();
            timer.Start();
            foreach (var hiddenComponent in hiddenComponents)
                modelDoc2.Extension.SelectByID2(hiddenComponent, "COMPONENT", 0, 0,  0, true, 1, null, 0);
            timer.Stop();
            Console.WriteLine($"Time spent with ModelDocExtension::SelectByID2  { timer.Elapsed.TotalSeconds} seconds");

            modelDoc2.ClearSelection();
            timer = new Stopwatch();
            timer.Start();
            modelDoc2.Extension.MultiSelect2(hiddenComponentObjects, false, null);
            timer.Stop();
            Console.WriteLine($"Time spent with ModelDocExtension::MultiSelect2: { timer.Elapsed.TotalSeconds} seconds");

            modelDoc2.EditSuppress2();

            swApp.CommandInProgress = false;

            timer.Stop();


            Console.WriteLine(timer.Elapsed.ToString());

        }

#endregion
    }
}