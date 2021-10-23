using BlueByte.SOLIDWORKS.Extensions;
using BlueByte.SOLIDWORKS.Helpers;
using SolidWorks.Interop.sldworks;
using System;
using System.Linq;

namespace Tests
{
    internal class Program
    {
        #region Private Methods

        private static void Main(string[] args)
        {
            var fileName = @"C:\Users\alexe\Desktop\TestAssemlby.SLDASM";
            var manager = new SOLIDWORKSInstanceManager();
            var swApp = default(SldWorks);

            var openRet = default(Tuple<bool, ModelDoc2>);
            var processes = System.Diagnostics.Process.GetProcessesByName("sldworks");

            if (processes != null)
            {
                var process = processes.FirstOrDefault();
                if (process != null)
                {
                    swApp = manager.GetSOLIDWORKSInstanceFromProcessID(process.Id);
                    if (swApp != null)
                    {
                        swApp.Visible = true;
                    }
                    else return;
                }
                else
                {
                    swApp = manager.GetNewInstance("", 60);
                    swApp.Visible = true;
                }
            }

            string[] errors;
            string[] warnings;
            var swModel = swApp.ActiveDoc as ModelDoc2;
            if (swModel == null)
            {
                openRet = swApp.OpenDocument(fileName, out errors, out warnings, SolidWorks.Interop.swconst.swOpenDocOptions_e.swOpenDocOptions_Silent);
                Tests.Do(openRet.Item2);
            }
            else
            {
                Tests.Do(swModel);
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

        public static void Do(ModelDoc2 modelDoc2)
        {
            // create bounding box
            //var data = modelDoc2.CreateBoundingBox();
            //modelDoc2.CreateRefPlanesAroundBoundingBox(data);  
            double[] data = null;
            modelDoc2.CreateRefPlanesAroundBoundingBox(data);
        }

        #endregion
    }
}