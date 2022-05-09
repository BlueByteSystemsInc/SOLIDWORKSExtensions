namespace BlueByte.SOLIDWORKS.Extensions.Helpers
{

    /// <summary>
    /// Help functions for dealing with rgb values.
    /// </summary>
    public static class ColorHelper
    {
        /// <summary>
        /// Gets the colorref from RGB.
        /// </summary>
        /// <param name="red">red.</param>
        /// <param name="green">green.</param>
        /// <param name="blue">blue.</param>
        /// <returns>colorref</returns>
        public static int GetCOLORREFfromRGB(int red, int green, int blue)
        {
            var rgb = red + (green * 256) + (blue * 65536);
            return rgb;
        }

        /// <summary>
        /// Gets the rgb values from colorref.
        /// </summary>
        /// <param name="value">value.</param>
        /// <returns></returns>
        public static int[] GetRGBfromCOLORREF(int value)
        {
          var color =  System.Drawing.ColorTranslator.FromWin32(value);


            return new int[] { color.R, color.G, color.B };
        }

    }
}
