using BlueByte.SOLIDWORKS.Extensions;
using BlueByte.SOLIDWORKS.Helpers;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tests
{
    class Program
    {
        static void Main(string[] args)
        {

            var fileName = @"C:\Users\jlili\Desktop\34492454\34492454.sldasm";
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

                        string[] errors;
                        string[] warnings;
                        if (string.IsNullOrWhiteSpace(fileName) == false)
                            openRet = swApp.OpenDocument(fileName, out errors, out warnings, SolidWorks.Interop.swconst.swOpenDocOptions_e.swOpenDocOptions_Silent);
                        Tests.Do(openRet.Item2);
                    }

                }
                else
                {
                    swApp = manager.GetNewInstance("", 60);
                    swApp.Visible = true;

                    string[] errors;
                    string[] warnings;
                    if (string.IsNullOrWhiteSpace(fileName) == false)
                        openRet = swApp.OpenDocument(fileName, out errors, out warnings, SolidWorks.Interop.swconst.swOpenDocOptions_e.swOpenDocOptions_Silent);

                    Tests.Do(openRet.Item2);
                }
            }
            else
            {

                swApp = manager.GetNewInstance("", 60);
                swApp.Visible = true;
                string[] errors;
                string[] warnings;
                if (string.IsNullOrWhiteSpace(fileName) == false)
                    openRet  = swApp.OpenDocument(fileName, out errors, out warnings, SolidWorks.Interop.swconst.swOpenDocOptions_e.swOpenDocOptions_Silent);

                Tests.Do(openRet.Item2);
            }

        }




    }


    public class Tests
    {
        public static void Do(SldWorks swApp)
        {

        }
        public static void Do(ModelDoc2 modelDoc2)
        {
            // create bounding box
            var data = modelDoc2.CreateBoundingBox();
            modelDoc2.CreateRefPlanesAroundBoundingBox(data);
        }
    }
}
