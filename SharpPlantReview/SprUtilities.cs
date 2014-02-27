//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Collections.Generic;
using System.Drawing;
using Microsoft.Win32;

namespace SharpPlant.SharpPlantReview
{
    public static class SprUtilities
    {
        /// <summary>
        ///     Returns a 24-bit color integer.
        /// </summary>
        /// <param name="rgbColor">The System.Drawing.Color to be converted.</param>
        /// <returns></returns>
        public static int Get0Bgr(Color rgbColor)
        {
            // Return a zero-alpha 24-bit BGR color integer
            return 0 + (rgbColor.B << 0x10) + (rgbColor.G << 0x8) + rgbColor.R;
        }

        /// <summary>
        ///     Returns a fully opaque (Alpha 255) color from a 0BGR format.
        /// </summary>
        /// <param name="bgrColor">The 0BGR integer to be converted.</param>
        /// <returns></returns>
        public static Color From0Bgr(int bgrColor)
        {
            // Get the color bytes
            var bytes = BitConverter.GetBytes(bgrColor);

            // Return the color from the byte array
            return Color.FromArgb(bytes[0], bytes[1], bytes[2]);
        }

        /// <summary>
        ///     Writes the tag properties to the default tag registry settings used by SmartPlant Review.
        /// </summary>
        /// <param name="tag">The tag used to determine settings values.</param>
        public static void SetTagRegistry(SprTag tag)
        {
            try
            {
                // Open the SmartPlant Review tag registry
                using (var regKey = Registry.CurrentUser.OpenSubKey(SprConstants.SprTagRegistryPath, true))
                {
                    if (regKey == null)
                    {
                        Registry.CurrentUser.CreateSubKey(SprConstants.SprTagRegistryPath);
                        SetTagRegistry(tag);
                        return;
                    }

                    var tagAtts = new List<Tuple<string, object, RegistryValueKind>>
                    {
                        new Tuple<string, object, RegistryValueKind>("BackgndColor", Get0Bgr(tag.BackgroundColor), RegistryValueKind.DWord),
                        new Tuple<string, object, RegistryValueKind>("ComputerName", tag.ComputerName, RegistryValueKind.String),
                        new Tuple<string, object, RegistryValueKind>("Creator", tag.Creator, RegistryValueKind.String),
                        new Tuple<string, object, RegistryValueKind>("Discipline", tag.Discipline, RegistryValueKind.String),
                        new Tuple<string, object, RegistryValueKind>("LeaderColor", Get0Bgr(tag.LeaderColor), RegistryValueKind.DWord),
                        new Tuple<string, object, RegistryValueKind>("Status", tag.Status, RegistryValueKind.String),
                        new Tuple<string, object, RegistryValueKind>("TextColor", Get0Bgr(tag.TextColor), RegistryValueKind.DWord)
                    };

                    foreach (var att in tagAtts)
                    {
                        var valString = string.Format("Default{0}", att.Item1);
                        regKey.SetValue(valString, att.Item2, att.Item3);
                    }
                }
            }
            catch (System.Security.SecurityException)
            {
                throw new SprException("Access to the SmartPlant Review registry denied.\n{0}",
                                       SprConstants.SprTagRegistryPath);
            }
        }

        /// <summary>
        ///     Clears the default tag settings from the registry.
        /// </summary>
        public static void ClearTagRegistry()
        {
            try
            {
                // Open the SmartPlant Review tag registry
                using (var regKey = Registry.CurrentUser.OpenSubKey(SprConstants.SprTagRegistryPath, true))
                {
                    if (regKey == null)
                        return;

                    var values = regKey.GetValueNames();
                    foreach (var t in values)
                        regKey.DeleteValue(t);
                }
            }
            catch (System.Security.SecurityException)
            {
                throw new SprException("Access to the SmartPlant Review registry denied.\n{0}",
                                       SprConstants.SprTagRegistryPath);
            }
        }

        /// <summary>
        ///     Checks for errors from a status returned by a DrApi function.
        /// </summary>
        /// <param name="errorStatus">The integer status returned by a DrApi function.</param>
        public static SprException GetError(int errorStatus)
        {
            var sprApp = SprApplication.ActiveApplication;
            if (sprApp == null || !sprApp.IsConnected)
                throw SprExceptions.SprNotConnected;

            if (errorStatus == 0)
                return null;

            string errorString;
            sprApp.DrApi.ErrorString(errorStatus, out errorString);
            if (sprApp.SprStatus != 0)
                throw sprApp.SprException;

            return new SprException(errorString);
        }
    }
}