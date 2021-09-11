using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BlueByte.SOLIDWORKS.Helpers
{
    public static class ModelDocHelper
    {

        /// <summary>
        /// Gets the active configuration material values.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <returns></returns>
        public static double[] GetActiveConfigurationMaterialValues(this ModelDoc2 doc)
        {
            var configuration = doc.GetActiveConfiguration() as Configuration;

            var displayStates = configuration.GetDisplayStates() as string[];

            var materials = doc.Extension.GetRenderMaterials2((int)swDisplayStateOpts_e.swAllDisplayState, displayStates) as object[];

            var material = materials.Select(x => x as RenderMaterial).First();

            int red = material.PrimaryColor & 255;

            int green = material.PrimaryColor >> 8 & 255;

            int blue = material.PrimaryColor >> 16 & 255;



            double Ambient = 0;
            double Diffuse = 1;
            double Specular = 1;
            double Shininess = 1;
            double Transparency = 0;
            double Emission = 0;

            return new double[] { red / 255.0, green / 255.0, blue / 255.0, Ambient, Diffuse, Specular, Shininess, Transparency, Emission };

        }
        /// <summary>
        /// Gets the color of the active configuration primary.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <returns></returns>
        public static double[] GetActiveConfigurationPrimaryColor(this ModelDoc2 doc)
        {
            var configuration = doc.GetActiveConfiguration() as Configuration;

            var displayStates = configuration.GetDisplayStates() as string[];

            var materials = doc.Extension.GetRenderMaterials2((int)swDisplayStateOpts_e.swAllDisplayState, displayStates) as object[];

            var nColor = materials.Select(x => x as RenderMaterial).First().PrimaryColor;

            int red = nColor & 255;

            int green = nColor >> 8 & 255;

            int blue = nColor >> 16 & 255;

            return new double[] { red, green, blue };

        }
        /// <summary>
        /// Gets the color of the active configuration primary reference.
        /// </summary>
        /// <param name="doc">The document.</param>
        /// <returns></returns>
        public static int GetActiveConfigurationPrimaryRefColor(this ModelDoc2 doc)
        {
            var configuration = doc.GetActiveConfiguration() as Configuration;

            var displayStates = configuration.GetDisplayStates() as string[];

            var materials = doc.Extension.GetRenderMaterials2((int)swDisplayStateOpts_e.swAllDisplayState, displayStates) as object[];
            var nColor = materials.Select(x => x as RenderMaterial).First().PrimaryColor;
            return nColor;


        }

        /// <summary>
        /// Sets the material color for the specified configuration.
        /// </summary>
        /// <param name="partDocModelDoc">Model document.</param>
        /// <param name="configurationName">Name of the configuration.</param>
        /// <param name="materialValues">The material values.</param>
        /// <exception cref="ArgumentNullException">
        /// partDocModelDoc
        /// or
        /// configurationName
        /// or
        /// materialValues
        /// </exception>
        /// <exception cref="Exception">
        /// Document type mismatch.
        /// or
        /// or
        /// Specified configuration is not active.
        /// </exception>
        public static void SetColor(this ModelDoc2 partDocModelDoc, string configurationName, double[] materialValues)
        {
            if (partDocModelDoc == null)
                throw new ArgumentNullException(nameof(partDocModelDoc));

            if (string.IsNullOrWhiteSpace(configurationName))
                throw new ArgumentNullException(nameof(configurationName));
            if (partDocModelDoc.GetType() != (int)swDocumentTypes_e.swDocPART)
                throw new Exception($"Document type mismatch.", new Exception($"{partDocModelDoc.GetTitle()} is not a part document."));
            var configuration = partDocModelDoc.GetActiveConfiguration() as Configuration;

            var displayStates = configuration.GetDisplayStates() as string[];


            partDocModelDoc.Extension.SetMaterialPropertyValues(materialValues, (int)swInConfigurationOpts_e.swThisConfiguration, new string[] { configuration.Name });
            partDocModelDoc.Extension.SetMaterialPropertyValues(materialValues, (int)swInConfigurationOpts_e.swThisConfiguration, new string[] { configuration.Name });

        }

        /// <summary>
        /// Creates the wrap assembly from.
        /// </summary>
        /// <param name="swApp">The SOLIDWORKS application.</param>
        /// <param name="swModelDoc">The SOLIDWORKS model document.</param>
        /// <param name="createBoundingBoxPlanes">if set to <c>true</c> [create bounding box planes].</param>
        /// <returns></returns>
        /// <exception cref="System.ArgumentNullException">
        /// swApp
        /// or
        /// swModelDoc
        /// </exception>
        /// <exception cref="System.Exception">Failed to add component.</exception>
        public static ModelDoc2 CreateWrapAssemblyFrom(this SldWorks swApp, ModelDoc2 swModelDoc, bool createBoundingBoxPlanes = false)
        {

            if (swApp == null)
                throw new System.ArgumentNullException(nameof(swApp));

            if (swModelDoc == null)
                throw new System.ArgumentNullException(nameof(swModelDoc));


            var modelPath = string.IsNullOrWhiteSpace(swModelDoc.GetPathName()) ? swModelDoc.GetTitle() : swModelDoc.GetPathName();


            var doc = swApp.NewAssembly() as ModelDoc2;
            var docAssembly = doc as AssemblyDoc;

            var addComponentRet = docAssembly.AddComponent(modelPath, 0, 0, 0);

            if (addComponentRet == false)
            {
                swApp.QuitDoc(doc.GetTitle());
                throw new System.Exception($"Failed to add component", new System.Exception($"Failed to add {swModelDoc.GetTitle()}"));
            }

            return doc;

        }

        /// <summary>
        /// Creates a bounding box feature and returns the Width, Thickness and Length in meters.
        /// </summary>
        /// <param name="swModelDoc">The sw model document.</param>
        /// <returns></returns>
        public static double[] CreateBoundingBox(this ModelDoc2 swModelDoc)
        {
            var boolstatus = swModelDoc.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swViewDispGlobalBBox, true);

            Feature BoundingBox;
            int longstatus = 0;

            BoundingBox = swModelDoc.FeatureManager.InsertGlobalBoundingBox((int)swGlobalBoundingBoxFitOptions_e.swBoundingBoxType_BestFit, true, false, out longstatus) as Feature;

            boolstatus = swModelDoc.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swViewDispGlobalBBox, false);

            string value;
            string resolvedvalue;

            var ar = new List<double>();


            swModelDoc.Extension.CustomPropertyManager["Default"].Get2("Total Bounding Box Width", out value, out resolvedvalue);
            ar.Add(double.Parse(resolvedvalue) / 1000.0);

            swModelDoc.Extension.CustomPropertyManager["Default"].Get2("Total Bounding Box Thickness", out value, out resolvedvalue);
            ar.Add(double.Parse(resolvedvalue) / 1000.0);

            swModelDoc.Extension.CustomPropertyManager["Default"].Get2("Total Bounding Box Length", out value, out resolvedvalue);
            ar.Add(double.Parse(resolvedvalue) / 1000.0);

            return ar.ToArray();

        }


        public enum VectorsSense_e
        {
            Same,
            Opposite,
        }

        public static bool areVectorsCollinear(MathVector V, MathVector U)
        {
            MathVector ret = V.Cross(U) as MathVector;
            double[] d = (double[])ret.ArrayData;
            if (Math.Round(d[0], 6) == 0 && Math.Round(d[1], 6) == 0 && Math.Round(d[2], 6) == 0)
                return true;
            else
                return false;
        }

        public static VectorsSense_e DoVectorsHaveSameSense(MathVector U, MathVector V)
        {

            if (U.Dot(V) > 0)
                return VectorsSense_e.Same;
            else
                return VectorsSense_e.Opposite;
        }

        public static T Get<T>(this ModelDoc2 modelDoc, string customPropertyName, out string value, string configuration = "Default")
        {
            if (modelDoc == null)
                throw new ArgumentNullException(nameof(modelDoc));



            string resolvedvalue = string.Empty;


            modelDoc.Extension.CustomPropertyManager[configuration].Get2(customPropertyName, out value, out resolvedvalue);


            T ret = default(T);

            try
            {
                if (typeof(T) == typeof(double))
                {
                    double v;
                    var parse = double.TryParse(resolvedvalue, out v);
                    return (T)Convert.ChangeType(v, typeof(T));
                }
                else if (typeof(T) == typeof(string))
                {
                    return (T)Convert.ChangeType(resolvedvalue, typeof(T));
                }
                else if (typeof(T) == typeof(bool))
                {
                    Boolean boo;
                    var parse = bool.TryParse(resolvedvalue, out boo);
                    return (T)Convert.ChangeType(boo, typeof(T));
                }



            }
            catch (Exception)
            {

            }


            return ret;
        }




        public static Feature[] CreateRefPlanesAroundBoundingBox(this ModelDoc2 swModel, double[] arr)
        {
            double Width = arr.First();
            double Thickness = arr[1];
            double length = arr.Last();




            swModel.SketchManager.Insert3DSketch(true);
            swModel.SketchManager.AddToDB = true;



            var firstSketchPoint = swModel.SketchManager.CreatePoint(Width * 0.5, Thickness * 0.5, length * 0.5);
            var secondSketchPoint = swModel.SketchManager.CreatePoint(Width * 0.5, Thickness * 0.5, -length * 0.5);
            var thirdSketchPoint = swModel.SketchManager.CreatePoint(Width * 0.5, -Thickness * 0.5, length * 0.5);
            var fourthSketchPoint = swModel.SketchManager.CreatePoint(Width * 0.5, -Thickness * 0.5, -length * 0.5);
            var fithSketchPoint = swModel.SketchManager.CreatePoint(-Width * 0.5, Thickness * 0.5, -length * 0.5);


            var sixthSketchPoint = swModel.SketchManager.CreatePoint(-Width * 0.5, Thickness * 0.5, length * 0.5);


            var seventhSketchPoint = swModel.SketchManager.CreatePoint(-Width * 0.5, -Thickness * 0.5, length * 0.5);
            var eightSketchPoint = swModel.SketchManager.CreatePoint(-Width * 0.5, -Thickness * 0.5, -length * 0.5);



            swModel.SketchManager.AddToDB = false;
            swModel.SketchManager.InsertSketch(true);

            var lastFeature = swModel.Extension.GetLastFeatureAdded();

            var sketchName = lastFeature.Name;

            swModel.ClearSelection();


            var featureL = new List<Feature>();

            // front plane

            seventhSketchPoint.Select2(false, 0);
            firstSketchPoint.Select2(true, 1);
            thirdSketchPoint.Select2(true, 2);

            var frontPlane = swModel.FeatureManager.InsertRefPlane((int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0, (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0, (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0) as Feature;
            frontPlane.Name = "front_plane";
            featureL.Add(frontPlane);

            // back plane 
            swModel.ClearSelection();

            secondSketchPoint.Select2(false, 0);
            fithSketchPoint.Select2(true, 1);
            fourthSketchPoint.Select2(true, 2);
            var backPlane = swModel.FeatureManager.InsertRefPlane((int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0, (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0, (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0) as Feature;
            backPlane.Name = "back_plane";
            featureL.Add(backPlane);

            // right plane
            swModel.ClearSelection();

            firstSketchPoint.Select2(false, 0);
            thirdSketchPoint.Select2(true, 1);
            fourthSketchPoint.Select2(true, 2);

            var rightPlane = swModel.FeatureManager.InsertRefPlane((int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0, (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0, (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0) as Feature;
            rightPlane.Name = "right_plane";
            featureL.Add(rightPlane);

            // left plane 
            sixthSketchPoint.Select2(false, 0);
            fithSketchPoint.Select2(true, 1);
            seventhSketchPoint.Select2(true, 2);

            var leftPlane = swModel.FeatureManager.InsertRefPlane((int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0, (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0, (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0) as Feature;
            leftPlane.Name = "left_plane";
            featureL.Add(leftPlane);


            // bottom plane 
            swModel.ClearSelection();
            seventhSketchPoint.Select2(false, 0);
            fourthSketchPoint.Select2(true, 1);
            thirdSketchPoint.Select2(true, 2);

            var bottomPlane = swModel.FeatureManager.InsertRefPlane((int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0, (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0, (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0) as Feature;
            bottomPlane.Name = "bottom_plane";
            featureL.Add(bottomPlane);

            // top plane
            swModel.ClearSelection();

            sixthSketchPoint.Select2(false, 0);
            firstSketchPoint.Select2(true, 1);
            secondSketchPoint.Select2(true, 2);

            var topPlane = swModel.FeatureManager.InsertRefPlane((int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0, (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0, (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident, 0) as Feature;
            topPlane.Name = "top_plane";
            featureL.Add(topPlane);


            leftPlane.Select2(false, 0);
            rightPlane.Select2(true, 1);

            var midPlane = swModel.FeatureManager.InsertRefPlane((int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_MidPlane, 0, (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_MidPlane, 0, 0, 0) as Feature;
            midPlane.Name = "mid_plane";
            featureL.Add(midPlane);


            return featureL.ToArray();

        }
        enum PointPosition
        {
            One,
            Two,
            Three,
            Four,
            Five,
            Six,
            Seven,
            Eight
        }





        /// <summary>
        /// Gets the SOLIDWORKS document type.
        /// </summary>
        /// <param name="path">Complete path of the SOLIDWORKS file.</param>
        /// <returns>Document type enum.</returns>
        public static swDocumentTypes_e GetSOLIDWORKSDocumentType(string path)
        {
            string extension = System.IO.Path.GetExtension(path);
            if (extension.ToLower().Contains("sldprt"))
            {
                return swDocumentTypes_e.swDocPART;
            }
            else if (extension.ToLower().Contains("sldasm"))
            {
                return swDocumentTypes_e.swDocASSEMBLY;
            }
            else if (extension.ToLower().Contains("slddrw"))
                return swDocumentTypes_e.swDocDRAWING;

            return swDocumentTypes_e.swDocNONE;
        }
    }

}
