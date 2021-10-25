using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
namespace BlueByte.SOLIDWORKS.Helpers
{

    public static class ReferenceGeometrySelectByID2Name_s
    {
        public const string Point = "DATUMPOINT";
        public const string Plane = "PLANE";
        public const string Axis = "AXIS";
    }
    public static class ComponentHelper
    {
        /// <summary>
        /// Gets tighest fit box.
        /// </summary>
        /// <param name="component">The component.</param>
        /// <returns></returns>
        public static double[] GetTighestFitBox(this Component2 component)
        {
            if (component == null)
                throw new ArgumentNullException(nameof(component));

            if (component.IsLoaded() == false)
                throw new Exception($"{nameof(GetTighestFitBox)} failed because {component.Name} is not loaded in memory");

            var bodies = new List<Body2>();
            TraverseComponentForComponents(component, (Component2 c) => {

                var cVisibility = (swComponentVisibilityState_e)c.Visible;

                switch (cVisibility)
                {
                    case swComponentVisibilityState_e.swComponentHidden:
                        return;
                    case swComponentVisibilityState_e.swComponentVisible:
                    case swComponentVisibilityState_e.swComponentUnknown:
                    default:
                        break;
                }

                var componentModel = c.GetModelDoc() as ModelDoc2;

                if (componentModel == null)
                    return;
                else
                {
                    var pathName = componentModel.GetPathName();
                    if (string.IsNullOrWhiteSpace(pathName))
                        return;

                    var arr_bodies = (c.GetBodies2((int)swBodyType_e.swAllBodies) as object[]);
                    if (arr_bodies != null)
                    {

                        var swBodies = arr_bodies.Cast<Body2>().ToList();
                        var transform = c.Transform2;
                        foreach (var swBody in swBodies)
                        {

                            var cCopy = swBody.Copy() as Body2;
                            var appliedTransform = cCopy.ApplyTransform(transform);
                            // math transform 
                            bodies.Add(cCopy);


                        }
                    }
                }

            });

            double maxX = 0;
            double minX = 0;
            double maxY = 0;
            double minY = 0;
            double maxZ = 0;
            double minZ = 0;

            double x = 0;
            double y = 0;
            double z = 0;



            for (int i = 0; i < bodies.Count; i++)
            {



                var swBody = bodies[i];



                swBody.GetExtremePoint(1, 0, 0, out x, out y, out z);
                if (i == 0 || x > maxX)
                    maxX = x;


                swBody.GetExtremePoint(-1, 0, 0, out x, out y, out z);
                if (i == 0 || x < minX)
                    minX = x;


                swBody.GetExtremePoint(0, 1, 0, out x, out y, out z);
                if (i == 0 || y > maxY)
                    maxY = y;


                swBody.GetExtremePoint(0, -1, 0, out x, out y, out z);
                if (i == 0 || y < minY)
                    minY = y;



                swBody.GetExtremePoint(0, 0, 1, out x, out y, out z);
                if (i == 0 || z > maxZ)
                    maxZ = z;

                swBody.GetExtremePoint(0, 0, -1, out x, out y, out z);
                if (i == 0 || z < minZ)
                    minZ = z;


            }

            var extremes = new double[6] { minX, minY, minZ, maxX, maxY, maxZ };


            return extremes;
        }

        public static void TraverseComponentForComponents(this Component2 component, Action<Component2> performAction)
        {

            var swComponent = component;
            if (swComponent != null)
            {
                performAction(swComponent);

                var components = swComponent.GetChildren() as object[];
                if (components != null)
                    foreach (var child in components)
                    {
                        TraverseComponentForComponents((child as Component2), performAction);

                    }

            }


        }


        public static void SetColor(this Component2 component, Color color)
        {
            SetColor(color, null,
                (m, o, c) => component.SetMaterialPropertyValues2(m, (int)o, c),
                (o, c) => component.RemoveMaterialProperty2((int)o, c));
        }
        static void GetColorScope(IComponent2 comp, out swInConfigurationOpts_e confOpts, out string[] confs)
        {
            confOpts = comp != null
                ? swInConfigurationOpts_e.swSpecifyConfiguration
                : swInConfigurationOpts_e.swThisConfiguration;
            confs = comp != null
                ? new string[] { comp.ReferencedConfiguration }
                : null;
        }
        static double[] ToMaterialProperties(Color color)
         => new double[]
         {
                color.R / 255d,
                color.G / 255d,
                color.B / 255d,
                1, 1, 0.5, 0.4,
                (255 - color.A) / 255d,
                0
         };
        internal static void SetColor(Color? color,
            IComponent2 ownerComp,
            Action<double[], swInConfigurationOpts_e, string[]> setColorAction,
            Action<swInConfigurationOpts_e, string[]> removeColorAction)
        {
            GetColorScope(ownerComp, out swInConfigurationOpts_e confOpts, out string[] confs);
            if (color.HasValue)
            {
                var matPrps = ToMaterialProperties(color.Value);
                setColorAction.Invoke(matPrps, confOpts, confs);
            }
            else
            {
                removeColorAction?.Invoke(confOpts, confs);
            }
        }
        /// <summary>
        /// Gets the type of the component's underlining model document.
        /// </summary>
        /// <param name="component">The component.</param>
        /// <returns>swDocumentTypes_e</returns>
        /// <exception cref="ArgumentNullException">component</exception>
        public static swDocumentTypes_e GetComponentUnderliningModelDocType(this Component2 component)
        {
            if (component == null)
                throw new ArgumentNullException(nameof(component));
            var pathName = component.GetPathName();
            return ModelDocHelper.GetSOLIDWORKSDocumentType(pathName);
        }

        //public static List<Body> GetBodies(this Feature feature)
        //{
        //    var subFeature = default(Feature);
        //    var swComponent = feature.GetSpecificFeature2() as Component2;
        //    if (swComponent != null)
        //    {
        //        var componentBodies = (swComponent.GetBodies2((int)swBodyType_e.swAllBodies) as object[]).Cast<Body>().ToList();
                
        //        subFeature = swComponent.FirstFeature();

        //        while (subFeature != null)
        //        {
        //            GetModelDocs(subFeature);
        //            subFeature = subFeature.GetNextFeature() as Feature;
        //        }
        //    }
        //}
    }
    public static class MateHelper
    {

        static string GetReferenceGeometrySelectByID2FromComponent(Component2 swComponent, string featureName, string subComponentPath = "")
        {
            /// finds a feature with a name that matches featureName
            /// and returns the SelectedByID2 type for that feature
            var GetFeatureTypeFromComponent = default(Func<Component2, string>);
            GetFeatureTypeFromComponent = (Component2 Component) =>
            {
                Feature feature = Component.FirstFeature();
                while (feature != null)
                {
                    Feature subfeature = feature.GetFirstSubFeature() as Feature;
                    while (subfeature != null)
                    {
                        if (subfeature.Name.ToLower() == featureName.ToLower())
                        {
                            string type = subfeature.GetTypeName2();
                            if (type == "RefPlane")
                                return ReferenceGeometrySelectByID2Name_s.Plane;
                            if (type == "RefAxis")
                                return ReferenceGeometrySelectByID2Name_s.Axis;
                            if (type == "RefPoint")
                                return ReferenceGeometrySelectByID2Name_s.Point;
                        }
                        subfeature = subfeature.GetNextSubFeature() as Feature;
                    }
                    if (feature.Name.ToLower() == featureName.ToLower())
                    {
                        string type = feature.GetTypeName2();
                        if (type == "RefPlane")
                            return ReferenceGeometrySelectByID2Name_s.Plane;
                        if (type == "RefAxis")
                            return ReferenceGeometrySelectByID2Name_s.Axis;
                        if (type == "RefPoint")
                            return ReferenceGeometrySelectByID2Name_s.Point;
                    }
                    feature = feature.GetNextFeature() as Feature;
                }
                return string.Empty;
            };
            var components = new List<Component2>();
            string componentPathName = swComponent.GetPathName();
            components.AddRange(new Component2[] { swComponent });
            foreach (var component in components)
            {
                var ret = GetFeatureTypeFromComponent(component);
                if (!string.IsNullOrWhiteSpace(ret))
                {
                    return ret;
                }
            }
            return string.Empty;
        }
        public static T Get<T>(this Component2 component, string customPropertyName, out string value, string configuration = "Default")
        {
            if (component == null)
                throw new ArgumentNullException(nameof(component));
            var modelDoc = component.GetModelDoc2() as ModelDoc2;
            if (modelDoc == null)
                throw new Exception("Could not get custom property.", new Exception($"[{component.GetSelectByIDString()}] suppression state is {((swComponentSuppressionState_e)component.GetSuppression2()).ToString()}"));
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

    }

}
