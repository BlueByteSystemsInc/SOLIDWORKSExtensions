using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Xml;

namespace BlueByte.SOLIDWORKS.Helpers
{
    public static class MaterialHelper
    {

        /// <summary>
        /// Gets all the available SOLIDWORKS materials.
        /// </summary>
        /// <param name="swApp">The SOLIDWORKS application.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException">swApp</exception>
        public static string[] GetSOLIDWORKSMaterials(this SldWorks swApp)
        {
            if (swApp == null)
                throw new ArgumentNullException(nameof(swApp));

            if (swApp.GetMaterialDatabaseCount() == 0)
               return  new string[] { };

            var materials = new List<string>();


            var mdatabaseObjects = swApp.GetMaterialDatabases() as object[];
            foreach (var mdatabaseObject in mdatabaseObjects)
            {
                string databasePath = mdatabaseObject as string;
                string[] materialsarray = GetMaterialNamesFromSOLIDWORKSMaterialDatabase(databasePath);
                materials.AddRange(materialsarray);
            }


            return materials.ToArray();

        }
   

        /// <summary>
        /// Gets the material names from a SOLIDWORKS material database.
        /// </summary>
        /// <param name="pathName">Name of the path.</param>
        /// <returns>Array of material names.</returns>
        /// <exception cref="ArgumentException">pathName is null or empty space.</exception>
        public static string[] GetMaterialNamesFromSOLIDWORKSMaterialDatabase(string pathName)
        {
            if (string.IsNullOrWhiteSpace(pathName))
                throw new ArgumentException("pathName is null or empty space.");

            List<string> mats = new List<string>();
            if (System.IO.File.Exists(pathName))
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.DtdProcessing = DtdProcessing.Parse;
                XmlReader reader = XmlReader.Create(pathName, settings);
                while (reader.Read())
                {
                    if (reader.Name == "material")
                    {
                        string m = reader.GetAttribute("name");
                        if (!string.IsNullOrEmpty(m))
                            mats.Add(m);
                    }
                }

            }

            return mats.ToArray();
        }
    
    
    
    }
}
