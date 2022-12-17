using BlueByte.SOLIDWORKS.Extensions.Enums;
using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static BlueByte.SOLIDWORKS.Extensions.SOLIDWORKSInstanceManager;

namespace BlueByte.SOLIDWORKS.Extensions
{
    public enum swAdvancedOpenDocOptions_e
    {
        swOpenDocOptions_Silent = 1,
        swOpenDocOptions_ReadOnly = 2,
        swOpenDocOptions_ViewOnly = 4,
        swOpenDocOptions_RapidDraft = 8,
        swOpenDocOptions_LoadModel = 16,
        swOpenDocOptions_AutoMissingConfig = 32,
        swOpenDocOptions_OverrideDefaultLoadLightweight = 64,
        swOpenDocOptions_LoadLightweight = 128,
        swOpenDocOptions_DontLoadHiddenComponents = 256,
        swOpenDocOptions_LoadExternalReferencesInMemory = 512,
        swOpenDocOptions_OpenDetailingMode = 1024,
        swOpenDocOptions_QuickView = -1,
    }


    public static class Extension
    {

        #region extension methods for ModelDoc


        public static string[] GetHiddenComponents(this ModelDoc2 swModelDoc)
        {

            if (swModelDoc == null)
                throw new ArgumentNullException(nameof(swModelDoc));

            if (swModelDoc.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                throw new Exception($"document is not assembly.");


            var assembly = swModelDoc as AssemblyDoc;

            if (assembly.GetComponentCount(true) == 0)
                throw new Exception("Assembly no components");

            var components = assembly.GetComponents(false) as object[];

            var swComponents = components.Select(x => x as Component2);

            var swPartComponents = swComponents.Where(x => x.IGetChildrenCount() == 0).ToArray();

            var swPartsComponentSelectionStrings = swPartComponents.Select(x => x.GetSelectByIDString()).ToArray();

          
            var ret = default(string[]);
            var retList = new List<string>();
          
            
            retList.AddRange(swPartsComponentSelectionStrings);


            swModelDoc.ViewZoomtofit2();


            if (swModelDoc.GetModelViewCount() == 0)
                throw new Exception($"{swModelDoc.GetTitle()} has no model views.");

            // get all standards views 
            var views = (object[])swModelDoc.GetModelViewNames();

            foreach (var view in views)
            {
                swModelDoc.ShowNamedView(view.ToString());
                var visibleComponents = assembly.GetVisibleComponentsInView();
                if (visibleComponents != null)
                {
                    var visibleComponentsInView = visibleComponents as object[];

                    foreach (var component in visibleComponentsInView)
                    {
                        var swComponent = component as Component2;
                        if (swComponent != null)
                        {
                            if (swComponent.IGetChildrenCount() == 0)
                            {
                                // only select parts 
                                var selectionStr = swComponent.GetSelectByIDString();
                                retList.Remove(selectionStr);
                                 
                            }
                        }
                    }
                }
            }

            return retList.ToArray();
        }
        public static Component2[] GetHiddenComponentObjects(this ModelDoc2 swModelDoc)
        {

            if (swModelDoc == null)
                throw new ArgumentNullException(nameof(swModelDoc));

            if (swModelDoc.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                throw new Exception($"document is not assembly.");


            var assembly = swModelDoc as AssemblyDoc;

            if (assembly.GetComponentCount(true) == 0)
                throw new Exception("Assembly no components");

            var components = assembly.GetComponents(false) as object[];

            var swComponents = components.Select(x => x as Component2);

            var swPartComponents = swComponents.Where(x => x.IGetChildrenCount() == 0).ToArray();

            var swPartsComponentSelectionStrings = swPartComponents.Select(x => x.GetSelectByIDString()).ToArray();


            var retList = new List<string>();
            var componentList = new List<Component2>();
            

            retList.AddRange(swPartsComponentSelectionStrings);
            componentList.AddRange(swPartComponents);

            swModelDoc.ViewZoomtofit2();


            if (swModelDoc.GetModelViewCount() == 0)
                throw new Exception($"{swModelDoc.GetTitle()} has no model views.");

            // get all standards views 
            var views = (object[])swModelDoc.GetModelViewNames();

            foreach (var view in views)
            {
                swModelDoc.ShowNamedView(view.ToString());
                var visibleComponents = assembly.GetVisibleComponentsInView();
                if (visibleComponents != null)
                {
                    var visibleComponentsInView = visibleComponents as object[];

                    foreach (var component in visibleComponentsInView)
                    {
                        var swComponent = component as Component2;
                        if (swComponent != null)
                        {
                            if (swComponent.IGetChildrenCount() == 0)
                            {
                                // only select parts 
                                var selectionStr = swComponent.GetSelectByIDString();
                                retList.Remove(selectionStr);
                                componentList.Remove(swComponent);
                            }
                        }
                    }
                }
            }

            return componentList.ToArray();
        }


        /// <summary>
        /// Removes all appearances. This method does not update graphics. Use <see cref="ModelDoc2.Rebuild3"/> to see changes.
        /// </summary>
        /// <param name="swModelDoc"></param>
        public static void RemoveAllAppearances(this ModelDoc2 document)
        {
            var displayStatesSettings = document.Extension.GetDisplayStateSetting((int)swDisplayStateOpts_e.swAllDisplayState);
            var renderMaterials = (object[])document.Extension.GetRenderMaterials2((int)swDisplayStateOpts_e.swAllDisplayState, displayStatesSettings.Names);
            if (renderMaterials != null)
            {
                foreach (var item in renderMaterials)
                {
                    var renderMaterial = item as RenderMaterial;
                    int Id1;
                    int Id2;
                    int[] DeleteId1 = new int[1];
                    int[] DeleteId2 = new int[1];
                    renderMaterial.GetMaterialIds(out Id1, out Id2);
                    DeleteId1[0] = Id1;
                    DeleteId2[0] = Id2;

                    var status = document.Extension.DeleteDisplayStateSpecificRenderMaterial(DeleteId1, DeleteId2);
                    document.Extension.UpdateRenderMaterialsInSceneGraph(true);
                }
            }
        }



        /// <summary>
        /// Opens an existing document in SOLIDWORKS. This methods supports non-SOLIDWORKS format as well.
        /// </summary>
        /// <param name="swApp">SOLIDWORKS application.</param>
        /// <param name="filename">complete pathname with filename and extension.</param>
        /// <param name="errors">An array that will contain descriptive errors. Empty if no errors. See description attributes from <seealso cref="swOpenError"/> and <seealso cref="swOpenWarning"/>.</param>
        /// <param name="warnings">An array will contain warning messages. Empty if no warnings.</param>
        /// <param name="openOptions">Open document options.</param>
        /// <param name="configurationName">Configuration name.</param>
        /// <returns>A tuple of true or false and a <see cref="ModelDoc2"/></returns>
        public static Tuple<bool,ModelDoc2> OpenDocument(this SldWorks swApp, string filename, out string[] errors, out string[] warnings, swAdvancedOpenDocOptions_e openOptions = swAdvancedOpenDocOptions_e.swOpenDocOptions_Silent, string configurationName = "")
        {
            warnings = new string[] { };
            errors = new string[] { };
            try
            {
                string[] nonSWCADExtensions = new string[] { ".step", ".stp", ".igs", ".stl" };
                swOpenError ErrorsEnum = default(swOpenError);
                swOpenWarning WarningsEnum = default(swOpenWarning);
                int Error = 0;
                int Warning = 0;
                swDocumentTypes_e type = default(swDocumentTypes_e);
                string extension = System.IO.Path.GetExtension(filename).ToLower();
                if (extension.Contains("prt"))
                    type = swDocumentTypes_e.swDocPART;
                else if (extension.Contains("asm"))
                    type = swDocumentTypes_e.swDocASSEMBLY;
                else if (extension.Contains("drw"))
                    type = swDocumentTypes_e.swDocDRAWING;
                else if (nonSWCADExtensions.Contains(extension.ToLower()))
                {
                    try
                    {
                        int err = 0;
                        var file = swApp.LoadFile4(filename, "", null, ref err) as ModelDoc2;
                        ErrorsEnum = (swOpenError)Error;
                        errors = GetDescriptions(GetFlags(ErrorsEnum));
                        if (errors.Length == 0)
                            errors = new string[] {  };
                        if (file == null)
                            return new Tuple<bool, ModelDoc2>(false, null);
                        else
                            return new Tuple<bool, ModelDoc2>(false, null);
                    }
                    catch (COMException e)
                    {
                        return new Tuple<bool, ModelDoc2>(false, null);
                    }
                }
                else
                    throw new Exception($"Cannot open file in SOLIDWORKS. Unknown extension [{System.IO.Path.GetFileName(filename)}].");


                ModelDoc2 Document = default(ModelDoc2);
                
                if (openOptions == swAdvancedOpenDocOptions_e.swOpenDocOptions_QuickView)
                {
                    var spec = swApp.GetOpenDocSpec(filename) as IDocumentSpecification;
                    spec.Selective = true;
                    spec.ConfigurationName = configurationName;
                    spec.Silent = true;
                    spec.DocumentType = (int)type;
                    Document = swApp.OpenDoc7(spec);

                   

                    Error = spec.Error;
                    Warning = spec.Warning;

                }
                else 

                 Document = swApp.OpenDoc6(filename, (int)type, (int)openOptions, configurationName, ref Error, ref Warning) as ModelDoc2;
                ErrorsEnum = (swOpenError)Error;
                WarningsEnum = (swOpenWarning)Warning;
                errors = GetDescriptions(GetFlags(ErrorsEnum));
                warnings = GetDescriptions(GetFlags(WarningsEnum));

      
                if (Error != 0 && Warning != 0)
                {
                    if (Document == null)
                        return new Tuple<bool, ModelDoc2>(false, null);
                }
                else if (Error != 0 && Warning == 0)
                {
                    if (Document == null)
                        return new Tuple<bool, ModelDoc2>(false, null);


                }
                else if (Error == 0 && Warning != 0)
                {
                    if (Document == null)
                        return new Tuple<bool, ModelDoc2>(false, null);
                    else
                        return new Tuple<bool, ModelDoc2>(true, Document);

                }

                if (Document == null)
                {
                    return new Tuple<bool, ModelDoc2>(false, null);
                }

                return new Tuple<bool, ModelDoc2>(true, Document);
            }
            catch (COMException e)
            {
                errors = new string[] { e.Message };



                return new Tuple<bool, ModelDoc2>(false, null);
            }
            catch (Exception ex)
            {
                errors = new string[] { ex.Message };

                return new Tuple<bool, ModelDoc2>(false, null);
            }
        }


        /// <summary>
        /// Returns whether a configuration exists or not
        /// </summary>
        /// <param name="modelDoc"></param>
        /// <param name="configurationName">configuration name</param>
        /// <returns>true if exists, false if not</returns>
        public static bool DoesConfigurationExist(this ModelDoc2 modelDoc, string configurationName)
        {
            if (configurationName == null) throw new ArgumentNullException("configurationName");

            if(modelDoc.GetConfigurationCount() > 0)
            {
                var configurations = (object[])modelDoc.GetConfigurationNames();
                for (int i = 0; i < configurations.Length; i++)
                {
                    if (configurations[i].ToString().ToLower().Trim() == configurationName.ToLower().Trim()) return true;
                }
            }

            return false;
        }



         static IEnumerable<Enum> GetFlags(Enum input)
        {
            foreach (Enum value in Enum.GetValues(input.GetType()))
                if (input.HasFlag(value))
                    yield return value;
        }
         static string[] GetDescriptions(IEnumerable<Enum> enums)
        {
            var list = new List<string>();
            foreach (var e in enums)
            {
                string error = GetEnumDescription(e);
                if (error != GetEnumDescription(swOpenError.swNoError) && error != GetEnumDescription(swOpenWarning.swNoError))
                    list.Add(error);
            }
            return list.ToArray();
        }
        internal static string GetEnumDescription(Enum value)
        {
            FieldInfo fi = value.GetType().GetField(value.ToString());

            DescriptionAttribute[] attributes =
                (DescriptionAttribute[])fi.GetCustomAttributes(
                typeof(DescriptionAttribute),
                false);

            if (attributes != null &&
                attributes.Length > 0)
                return attributes[0].Description;
            else
                return value.ToString();
        }

        #endregion 

        #region methods to get sldworks 

        /// <summary>
        /// Returns the SOLIDWORKS installation directory for the specified year (if it exists).
        /// </summary>
        /// <param name="Year">Year</param>
        /// <returns><see cref="DirectoryInfo"/> object, null if failed.</returns>
        public static DirectoryInfo GetSOLIDWORKSInstallationDirectory(int Year)
        {
            using (RegistryKey hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            {
                using (RegistryKey key = hklm.OpenSubKey(@"SOFTWARE\SolidWorks\SOLIDWORKS " + Year + @"\Setup"))
                {
                    if (key == null)
                        return null;
                    else
                    {
                        return new DirectoryInfo(key.GetValue("SolidWorks Folder") as string);
                    }
                }
            }
        }

        /// <summary>
        /// Creates a new instance of the SOLIDWORKS application using <see cref="Process.Start()"/>.
        /// </summary>
        /// <param name="timeoutSec"></param>
        /// <param name="suppressDialog">True to suppress SOLIDWORKS dialogs.</param>
        /// <returns>Pointer to the new instance of SOLIDWORKS.</returns>
        /// <exception cref="TimeoutException">Thrown if method times out.</exception>
        public static SldWorks CreateSldWorks(string commandlineParameters = "", Year_e _year =  Year_e.Latest, int timeoutSec = 30)
        {
            int[] years = ReleaseYears();
            if (years.Length == 0)
                throw new Exception($"SOLIDWORKS is not installed on this computer [{System.Environment.MachineName}].");
            Array.Sort(years);
            
            int year = years.Last();

            if (_year != (int)Year_e.Latest)
            {
                if (years.Contains((int)_year) == false)
                    throw new Exception($"Could not find installation directory for SOLIDWORKS. [{_year}].");

                year = (int)_year;
            }
            
            var installationDirectory = GetSOLIDWORKSInstallationDirectory(year);
            if (installationDirectory == null)
                throw new Exception($"Could not find installation directory for SOLIDWORKS. Year = [{year}].");

            string appPath = installationDirectory.FullName;
            var timeout = TimeSpan.FromSeconds(timeoutSec);
            var startTime = DateTime.Now;
            string args = string.IsNullOrWhiteSpace(commandlineParameters) ? "/r" : commandlineParameters;
            var prc = Process.Start(appPath + "sldworks.exe", args);
            SldWorks app = null;
            while (app == null)
            {
                if (DateTime.Now - startTime > timeout)
                {
                    if (prc.Id != 0)
                        if (prc.HasExited == false)
                    prc.Kill();

                    throw new TimeoutException($"Could not create a new SOLIDWORKS process within {timeoutSec} seconds.");
                }

                app = GetSwAppFromProcess(prc.Id);
            }


            
            return app;
        }


        /// <summary>
        /// Returns an array of all installed SOLIDWORKS release years.
        /// </summary>
        /// <returns>Array of integers.</returns>
        public static int[] ReleaseYears()
        {
            var solidworksKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\SolidWorks");
            var names = solidworksKey.GetSubKeyNames();
            var years = new List<int>();
            if (names == null)
                return years.ToArray();


            var regex = new Regex(@"^solidworks ([\d]{4})$", RegexOptions.IgnoreCase);
            foreach (var name in names)
            {
                if (regex.IsMatch(name))
                {
                    int year = int.MinValue;
                    var match = regex.Match(name);
                    var capture = match.Groups[1].Value;
                    var ret = int.TryParse(capture, out year);
                    if (ret)
                        years.Add(year);
                }
            }

            return years.ToArray();
        }




        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        private static SldWorks GetSwAppFromProcess(int processId)
        {
            var monikerName = "SolidWorks_PID_" + processId.ToString();

            IBindCtx context = null;
            IRunningObjectTable rot = null;
            IEnumMoniker monikers = null;

            try
            {
                CreateBindCtx(0, out context);

                context.GetRunningObjectTable(out rot);
                rot.EnumRunning(out monikers);

                var moniker = new IMoniker[1];

                while (monikers.Next(1, moniker, IntPtr.Zero) == 0)
                {
                    var curMoniker = moniker.First();

                    string name = null;

                    if (curMoniker != null)
                    {
                        try
                        {
                            curMoniker.GetDisplayName(context, null, out name);
                        }
                        catch (UnauthorizedAccessException)
                        {
                        }
                    }

                    if (string.Equals(monikerName,
                        name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        object app;
                        rot.GetObject(curMoniker, out app);
                        return (SldWorks)app;
                    }
                }
            }
            finally
            {
                if (monikers != null)
                {
                    Marshal.ReleaseComObject(monikers);
                }

                if (rot != null)
                {
                    Marshal.ReleaseComObject(rot);
                }

                if (context != null)
                {
                    Marshal.ReleaseComObject(context);
                }
            }

            return null;
        }




        /// <summary>
        /// Converts a release year to the SOLIDWORKS revision number.
        /// </summary>
        /// <param name="releaseYear">Release year.</param>
        /// <returns>SOLIDWORKS revision number</returns>
        /// <remarks>This method produces correct results for revision number from the year 2003 and newer.</remarks>
        public static int ConvertYearToSWRevisionNumber(int releaseYear)
        {
            return releaseYear - 1992;
        }

        /// <summary>
        /// Converts the SOLIDWORKS revision number to a release year.
        /// </summary>
        /// <param name="revNumber">Revision number.</param>
        /// <returns>Release year</returns>
        /// <remarks><ul>
        /// <li>This method produces correct results for revision number from the year 2003 and newer.</li>
        /// <li>Method return -1 if it fails.</li>
        /// </ul></remarks>
        public static int ConvertSWRevisionNumberToYear(string revNumber)
        {
            try
            {
                int spN = int.Parse(revNumber.Split('.').First());
                return 1992 + spN;
            }
            catch (Exception)
            {
                return -1;
            }

        }
        /// <summary>
        /// Converts the SOLIDWORKS revision number to a release year.
        /// </summary>
        /// <param name="revNumber">Revision number.</param>
        /// <returns>Release year</returns>
        /// <remarks><ul>
        /// <li>This method produces correct results for revision number from the year 2003 and newer.</li>
        /// <li>Method return -1 if it fails.</li>
        /// </ul></remarks>
        public static int ConvertSWRevisionNumberToYear(int revNumber)
        {

            return 1992 + revNumber;


        }



        #endregion 


    }
}
