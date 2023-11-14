using BlueByte.SOLIDWORKS.Extensions;
using BlueByte.SOLIDWORKS.Extensions.Enums;
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
            GetSheetMetalFeatures_Test();
            //ExportFlatPattern_Test();

        }

        
        
        private static void GetSheetMetalFeatures_Test()
        {
            var instanceManage = new SOLIDWORKSInstanceManager();
            var swApp = instanceManage.GetNewInstance();


            swApp.Visible = true;


            swApp.MaximizeWindow();


            string[] warnings;
            string[] errors;

            var fileName = @"C:\SOLIDWORKSPDM\Bluebyte\API\Sandbox\Grill Assembly\range_burner_insert_&.SLDPRT";


            var doc = swApp.OpenDocument(fileName, out errors, out warnings);

            var model = doc.Item2;


            var features = model.FeatureManager.GetFeatures(false) as object[];

            foreach (var feature in features)
            {
                var swFeature = feature as Feature;

                SheetMetalFeatureType_e t;
                var isSheetMetal = swFeature.IsSheetMetal(out t);

                Console.WriteLine($"{swFeature.Name.PadRight(50)} - IsSheetMetal {isSheetMetal} - {t.ToString()} ");
            }


            swApp.CloseAllDocuments(false);
            swApp.ExitApp();

            instanceManage.ReleaseInstance(swApp);
        }

        private static void ExportFlatPattern_Test()
        {
            var instanceManage = new SOLIDWORKSInstanceManager();
            var swApp = instanceManage.GetNewInstance();


            swApp.Visible = true;


            swApp.MaximizeWindow();


            string[] warnings;
            string[] errors;

            var fileName = @"C:\Users\jlili\Downloads\26050567.SLDPRT";


            var doc = swApp.OpenDocument(fileName, out errors, out warnings);

            var model = doc.Item2;

            var ret = model.ExportFlatPattern(SheetMetalOptions.SheetMetalOptions_ExportGeometry, System.IO.Path.ChangeExtension(fileName, "dxf"));


            swApp.CloseAllDocuments(false);
            swApp.ExitApp();

            instanceManage.ReleaseInstance(swApp);
        }



    }
 
}