//
//  Copyright © 2013 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using Microsoft.Win32;

namespace SharpPlant.SharpPlantReview
{
    public static class SprUtilities
    {

        /// <summary>
        ///     Data template for creating a new SprTag object.
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, object> TagTemplate()
        {
            return new Dictionary<string, object>()
            {
                {"tag_unique_id", 0},
                {"tag_size", 0},
                {"linkage_id_0", 0},
                {"linkage_id_1", 0},
                {"linkage_id_2", 0},
                {"linkage_id_3", 0},
                {"tag_text", string.Empty},
                {"number_color", 0},
                {"backgnd_color", 0},
                {"leader_color", 0},
                {"discipline", string.Empty},
                {"creator", string.Empty},
                {"computer_name", string.Empty},
                {"status", string.Empty}
            };
        }

        /// <summary>
        ///     Data template for creatring a new SprAnnotation object.
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, object> AnnotationTemplate()
        {
            return new Dictionary<string, object>()
            {
                {"id", 0},
                {"type_id", 0},
                {"bg_color", 0},
                {"line_color", 0},
                {"text_color", 0},
                {"text_string", string.Empty},
            };
        }

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
            // Open the SmartPlant Review tag registry
            const string regPath = @"Software\Intergraph\SmartPlant Review\Settings\Tags\";
            using (var regKey = Registry.CurrentUser.OpenSubKey(regPath, true))
            {
                if (regKey == null)
                {
                    Registry.CurrentUser.CreateSubKey(regPath);
                    SetTagRegistry(tag);
                    return;
                }

                // Create a three-variable list containing the subkey name, value and type.
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

                // Iterate through the attributes
                foreach (var att in tagAtts)
                {
                    // Get the subkey name (prefixed with "Default" per Spr formatting)
                    var valString = string.Format("Default{0}", att.Item1);

                    // Set the subkey values from the current tuple
                    regKey.SetValue(valString, att.Item2, att.Item3);
                }
            }
        }

        /// <summary>
        ///     Clears the default tag settings from the registry.
        /// </summary>
        public static void ClearTagRegistry()
        {
            // Open the SmartPlant Review tag registry
            const string regPath = @"Software\Intergraph\SmartPlant Review\Settings\Tags";
            using (var regKey = Registry.CurrentUser.OpenSubKey(regPath, true))
            {
                // Get the list of values
                var values = regKey.GetValueNames();

                // Iterate through the values
                foreach (var t in values)
                {
                    // Delete the current value
                    regKey.DeleteValue(t);
                }
            }
        }

        /// <summary>
        ///     Serializes a DataRow into a new SprTag.
        /// </summary>
        /// <param name="row">The DataRow containing the tag values.</param>
        /// <returns>The constructed SprTag.</returns>
        [Obsolete]
        public static SprTag BuildTagFromData(System.Data.DataRow row)
        {
            if (row == null)
                return null;

            var returnTag = new SprTag();

            // Add the values to the tag data dictionary
            returnTag.Row.ItemArray = row.ItemArray;

            return returnTag;
        }
		
		/// <summary>
        ///     Serializes a SprAnnotation from a qualified DataRow
        /// </summary>
        /// <param name="row">The DataRow containing the annotation values.</param>
        /// <returns>SprAnnotation</returns>
        public static SprAnnotation BuildAnnotationFromData(System.Data.DataRow row)
        {
            var returnAnnotation = new SprAnnotation();
            for (int i = 0; i < row.Table.Columns.Count; i++)
			{
                var key = row.Table.Columns[i].ColumnName;
                returnAnnotation.Data[key] = row.ItemArray[i];
			}

            return returnAnnotation;
        }

        /// <summary>
        ///     Checks for errors from a status returned by a DrApi function.
        /// </summary>
        /// <param name="status">The integer status returned by a DrApi function.</param>
        public static void ErrorCheck(int status)
        {
            var sprApp = SprApplication.ActiveApplication;
            if (sprApp == null)
                throw SprExceptions.SprNotConnected;

            if (status == 0)
                return;
            
            var errorString = string.Empty;
            sprApp.DrApi.ErrorString(status, out errorString);

            sprApp.LastError = errorString;
            throw new SprException(errorString);
        }
    }
}