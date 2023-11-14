using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BlueByte.SOLIDWORKS.Extensions.Enums;
namespace BlueByte.SOLIDWORKS.Extensions
{
  

    public static class FeatureHelper
    {

        public static bool IsSheetMetal(this Feature feature, out SheetMetalFeatureType_e type)
        {
            var featureType = feature.GetTypeName2();
            
            var names = Enum.GetNames(typeof(SheetMetalFeatureType_e));

            if (names.Contains(featureType))
            {
                type = (SheetMetalFeatureType_e)Enum.Parse(typeof(SheetMetalFeatureType_e), featureType); 
                return true; 
            }


            type = SheetMetalFeatureType_e.Unknown;
            return false; 

        }
    }
}
