using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;

namespace BlueByte.SOLIDWORKS.Helpers
{
    public static class SelectionHelper
    {
        /// <summary>
        /// Gets selected objects of type <see cref="T"/>.
        /// </summary>
        /// <typeparam name="T">SOLIDWORKS API type</typeparam>
        /// <param name="swModel">The SOLIDWORKS model document.</param>
        /// <param name="selectionMark">The selection mark.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException">swModel</exception>
        public static T[] GetSelectedObjects<T>(this ModelDoc2 swModel, int selectionMark = -1)
        {

            if (swModel == null)
                throw new ArgumentNullException(nameof(swModel));

                var list = new List<T>();
                var selectionMgr = swModel.SelectionManager as SelectionMgr;
                int selectedobjectsCount = selectionMgr.GetSelectedObjectCount2(selectionMark);
                if (selectedobjectsCount == 0)
                    return new T[] { };
                else
                {
                    for (int i = 1; i <= selectedobjectsCount; i++)
                    {

                    var obj = default(T);
                    try
                    {
                         obj = (T)selectionMgr.GetSelectedObject6(i, selectionMark);

                    }
                    catch (InvalidCastException)
                    {
                        obj = default(T);
                    }
                       
                     if (EqualityComparer<T>.Default.Equals(obj) == false)
                      list.Add(obj);

                    }
                }
                return list.ToArray();
            
        }
        /// <summary>
        /// Gets selected objects of type <see cref="T"/>.
        /// </summary>
        /// <typeparam name="T">SOLIDWORKS API type</typeparam>
        /// <param name="SelectionMgr">The selection manager</param>
        /// <param name="selectionMark">The selection mark.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException">swModel</exception>
        public static T[] GetSelectedObjects<T>(this SelectionMgr selectionMgr, int selectionMark = -1)
        {

            if (selectionMgr == null)
                throw new ArgumentNullException(nameof(selectionMgr));

                var list = new List<T>();
                int selectedobjectsCount = selectionMgr.GetSelectedObjectCount2(selectionMark);
                if (selectedobjectsCount == 0)
                    return new T[] { };
                else
                {
                    for (int i = 1; i <= selectedobjectsCount; i++)
                    {

                    var obj = default(T);
                    try
                    {
                         obj = (T)selectionMgr.GetSelectedObject6(i, selectionMark);

                    }
                    catch (InvalidCastException)
                    {
                        obj = default(T);
                    }
                       
                     if (EqualityComparer<T>.Default.Equals(obj) == false)
                      list.Add(obj);

                    }
                }
                return list.ToArray();
            
        }
    }
     

}
