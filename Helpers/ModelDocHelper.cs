using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BlueByte.SOLIDWORKS.Helpers
{

    /// <summary>
    /// Sheet metal options
    /// </summary>
    public enum SheetMetalOptions
        {
            SheetMetalOptions_ExportGeometry = 1,
            SheetMetalOptions_IncludeHiddenLines = 2,
            SheetMetalOptions_ExportBendLines = 4,
            SheetMetalOptions_IncludeSketches = 8,
            SheetMetalOptions_MergeCoPlanarFaces = 16,
            SheetMetalOptions_ExportLibraryFeatures = 32,
            SheetMetalOptions_ExportFormTools = 64,
            SheetMetalOptions_ExportBoundingBox = 4096
        }

    public static class ModelDocHelper
    {
        #region Public Enums

        public enum VectorsSense_e
        {
            Same,
            Opposite,
        }

        #endregion

        #region Public Methods



        /// <summary>
        /// Exports the flat pattern.
        /// </summary>
        /// <param name="model">The model.</param>
        /// <param name="exportOptions">The export options.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException">nameof(model)</exception>
        /// <exception cref="Exception">
        /// $"{model.GetTitle()} is not a part document.
        /// or
        /// $"{model.GetTitle()} is not a sheet metal part document.
        /// </exception>
        public static bool ExportFlatPattern(this ModelDoc2 model, SheetMetalOptions exportOptions, string fileName)
        {
            if (model == null)
                throw new ArgumentNullException(nameof(model)); 

            if (model.GetType() != (int)swDocumentTypes_e.swDocPART)
                throw new Exception($"{model.GetTitle()} is not a part document.");

            var partDoc = model as PartDoc;


            if (model.IsSheetMetal() == false)
            throw new Exception($"{model.GetTitle()} is not a sheet metal part document.");
            

            var sModelName = model.GetPathName();
            object varAlignment;
            double[] dataAlignment = new double[12];

            dataAlignment[0] = 0.0;
            dataAlignment[1] = 0.0;
            dataAlignment[2] = 0.0;
            dataAlignment[3] = 1.0;
            dataAlignment[4] = 0.0;
            dataAlignment[5] = 0.0;
            dataAlignment[6] = 0.0;
            dataAlignment[7] = 1.0;
            dataAlignment[8] = 0.0;
            dataAlignment[9] = 0.0;
            dataAlignment[10] = 0.0;
            dataAlignment[11] = 1.0;
            varAlignment = dataAlignment;
            return partDoc.ExportToDWG2(fileName, sModelName, (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal, true, varAlignment, false, false, (int)exportOptions, null);
        }


        /// <summary>
        /// Gets the reference document from drawing.
        /// </summary>
        /// <param name="modelDoc">The model document.</param>
        /// <returns></returns>
        /// <exception cref="System.Exception">
        /// </exception>
        public static ModelDoc2 GetReferenceDocumentFromDrawing(this ModelDoc2 modelDoc)
        {
           
                var drawing = modelDoc as DrawingDoc;

            if (drawing == null)
                throw new Exception($"{modelDoc.GetTitle()} is not a drawing document.");


            if (drawing.IsDetailingMode())
                throw new Exception($"{modelDoc.GetTitle()} is open in detailing mode. Could get referenced document to export flat pattern.");


            var referencedDocument = default(ModelDoc2);



                    var views = drawing.GetViews() as object[];
                    if (views != null)
                    {
                        var drawingViews = views.First() as object[];
                        foreach (var view in drawingViews)
                        {
                            var swView = view as SolidWorks.Interop.sldworks.View;
                            if (swView != null)
                            {
                                referencedDocument = swView.ReferencedDocument;
                        if (referencedDocument != null)
                            break;
                            }
                        }                        
                    }



            return referencedDocument;
        }

        /// <summary>
        /// Determines whether [is sheet metal] [the specified model].
        /// </summary>
        /// <param name="model">The model.</param>
        /// <returns>
        ///   <c>true</c> if [is sheet metal] [the specified model]; otherwise, <c>false</c>.
        /// </returns>
        /// <exception cref="ArgumentNullException">nameof(model)</exception>
        /// <exception cref="Exception">$"{model.GetTitle()} is not a part document.</exception>
        public static bool IsSheetMetal(this ModelDoc2 model)
        {
          
            var ret = false; 

               if (model == null)
                throw new ArgumentNullException(nameof(model)); 

            if (model.GetType() != (int)swDocumentTypes_e.swDocPART)
                throw new Exception($"{model.GetTitle()} is not a part document.");


            var partDoc = model as PartDoc;

            var isSheetMetal = true;

            var bodies = partDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);

            if (bodies != null)
            {
                var swBodies = (bodies as object[]).Cast<Body2>();
                if (swBodies != null)
                    foreach (var swBody in swBodies)
                    {
                        isSheetMetal = swBody.IsSheetMetal();
                        if (isSheetMetal)
                            break;
                    }

            }

            if (isSheetMetal)
                return true;

            var sheetMetal = (model.FeatureManager.GetFeatures(true) as object[]).Cast<Feature>().ToList().FirstOrDefault(i => i.GetTypeName2() == "SheetMetal");

            if (sheetMetal != null)
                isSheetMetal = true;

            ret = isSheetMetal;

            return ret; 
       
        }
            

        /// <summary>
        /// Converts to meters using the document's length unit.
        /// </summary>
        /// <param name="swModelDoc">The sw model document.</param>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        /// <exception cref="System.NotImplementedException">Conversion from {lengtUnit.ToString()} to meters not implemented by this method</exception>
        public static double ConvertToMeters(this ModelDoc2 swModelDoc, double value)
        {
            double factor = 1000.0;
            var ar = new List<double>();

            var lengtUnit = (swLengthUnit_e)swModelDoc.LengthUnit;
            switch (lengtUnit)
            {
                case swLengthUnit_e.swMM:
                    factor = 1 / 1000.0;
                    break;

                case swLengthUnit_e.swCM:
                    factor = 1 / 100.0;
                    break;

                case swLengthUnit_e.swMETER:
                    factor = 1;
                    break;

                case swLengthUnit_e.swINCHES:
                    factor = 1 / 39.37;
                    break;

                case swLengthUnit_e.swFEET:
                    factor = 1 / 3.281;
                    break;

                case swLengthUnit_e.swFEETINCHES:
                    factor = 1 / 3.281;
                    break;

                case swLengthUnit_e.swUIN:
                case swLengthUnit_e.swANGSTROM:
                case swLengthUnit_e.swNANOMETER:
                case swLengthUnit_e.swMICRON:
                case swLengthUnit_e.swMIL:
                default:
                    throw new NotImplementedException($"Conversion from {lengtUnit.ToString()} to meters not implemented by this method");
            }


            return value * factor;
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

        /// <summary>
        /// Creates a bounding box feature and returns the Width, Thickness and Length in meters.
        /// </summary>
        /// <param name="swModelDoc">The sw model document.</param>
        /// <returns></returns>
        public static double[] CreateBoundingBox(this ModelDoc2 swModelDoc)
        {
            string value;
            string resolvedvalue;
            double factor = 1000.0;
            var ar = new List<double>();

            var lengtUnit = (swLengthUnit_e)swModelDoc.LengthUnit;
            switch (lengtUnit)
            {
                case swLengthUnit_e.swMM:
                    factor = 1 / 1000.0;
                    break;

                case swLengthUnit_e.swCM:
                    factor = 1 / 100.0;
                    break;

                case swLengthUnit_e.swMETER:
                    factor = 1;
                    break;

                case swLengthUnit_e.swINCHES:
                    factor = 1 / 39.37;
                    break;

                case swLengthUnit_e.swFEET:
                    factor = 1 / 3.281;
                    break;

                case swLengthUnit_e.swFEETINCHES:
                    factor = 1 / 3.281;
                    break;

                case swLengthUnit_e.swUIN:
                case swLengthUnit_e.swANGSTROM:
                case swLengthUnit_e.swNANOMETER:
                case swLengthUnit_e.swMICRON:
                case swLengthUnit_e.swMIL:
                default:
                    throw new NotImplementedException($"Conversion from {lengtUnit.ToString()} to meters not implemented by this method");
            }

            var configurationName = swModelDoc.ConfigurationManager.ActiveConfiguration.Name;

            Feature BoundingBox;
            int longstatus = 0;

            var features = swModelDoc.FeatureManager.GetFeatures(true) as object[];
            var swfeatures = features.Cast<Feature>();
            var bbFeature = swfeatures.FirstOrDefault(x => x.GetTypeName() == "BoundingBox");
            if (bbFeature == null)
            {
                var boolstatus = swModelDoc.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swViewDispGlobalBBox, true);
                BoundingBox = swModelDoc.FeatureManager.InsertGlobalBoundingBox((int)swGlobalBoundingBoxFitOptions_e.swBoundingBoxType_BestFit, true, false, out longstatus) as Feature;
                boolstatus = swModelDoc.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swViewDispGlobalBBox, false);
            }

            swModelDoc.Extension.CustomPropertyManager[configurationName].Get2("Total Bounding Box Width", out value, out resolvedvalue);
            ar.Add(double.Parse(resolvedvalue) * factor);

            swModelDoc.Extension.CustomPropertyManager[configurationName].Get2("Total Bounding Box Thickness", out value, out resolvedvalue);
            ar.Add(double.Parse(resolvedvalue) * factor);

            swModelDoc.Extension.CustomPropertyManager[configurationName].Get2("Total Bounding Box Length", out value, out resolvedvalue);
            ar.Add(double.Parse(resolvedvalue) * factor);

            return ar.ToArray();
        }

        public static void TraverseFeatureForComponents(this Feature swFeature, Action<Component2> performAction)
        {
            var swSubFeature = default(Feature);

            var swComponent = swFeature.GetSpecificFeature2() as Component2;
            if (swComponent != null)
            {
                performAction(swComponent);

                swSubFeature = swComponent.FirstFeature();
                while (swSubFeature != null)
                {
                    TraverseFeatureForComponents(swSubFeature, performAction);
                    swSubFeature = swSubFeature.GetNextFeature() as Feature;
                }
            }


        }



        /// <summary>
        /// Creates top, bottom, right, left, front and back planes around the tightest fit bounding box of the component.
        /// </summary>
        /// <param name="swModel">The sw model.</param>
        /// <param name="component">The component.</param>
        /// <returns></returns>
        public static Feature[] CreateRefPlanesAroundBoundingBoxAroundComponent(this ModelDoc2 swModel, Component2 component)
        {
    
            double maxX = 0;
            double minX = 0;
            double maxY = 0;
            double minY = 0;
            double maxZ = 0;
            double minZ = 0;

            

            var box = component.GetTighestFitBox();



            minX = box[0];
            minY = box[1];
            minZ = box[2];
            maxX = box[3];
            maxY = box[4];
            maxZ = box[5];

            swModel.SketchManager.Insert3DSketch(true);
            swModel.SketchManager.AddToDB = true;

            var firstSketchPoint = swModel.SketchManager.CreatePoint(maxX, maxY, maxZ);
            var secondSketchPoint = swModel.SketchManager.CreatePoint(maxX, maxY,  minZ);
 
            var thirdSketchPoint = swModel.SketchManager.CreatePoint(maxX, minY,maxZ);
            var fourthSketchPoint = swModel.SketchManager.CreatePoint(maxX, minY, minZ);
            var fithSketchPoint = swModel.SketchManager.CreatePoint(minX, maxY, minZ);
            var sixthSketchPoint = swModel.SketchManager.CreatePoint(minX, maxY, maxZ);
            var seventhSketchPoint = swModel.SketchManager.CreatePoint(minX, minY, maxZ);
            var eightSketchPoint = swModel.SketchManager.CreatePoint(minX, minY, minZ);

            swModel.SketchManager.AddToDB = false;
            swModel.SketchManager.Insert3DSketch(true);

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

        #endregion
    }
}