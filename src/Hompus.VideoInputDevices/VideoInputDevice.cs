using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using System.Windows.Input;

namespace Hompus.VideoInputDevices
{
    /// <summary>
    /// A video input device that is detected in the system.
    /// </summary>
    public class VideoInputDevice
    {
        public string dataString { get => $"{_key.ToString()}"; }
        private FullName _key;
        public FullName Data { get => _key; }

        /// <summary>
        /// Initializes a new instance of the <see cref="VideoInputDevice"/> class.
        /// </summary>
        /// <param name="moniker">A moniker object.</param>
        public VideoInputDevice(IMoniker moniker)
        {
             _key = this.GetFriendlyName(moniker);
        }

        /// <summary>
        /// Get the name represented by the moniker.
        /// </summary>
        private FullName GetFriendlyName(IMoniker moniker)
        {
            object bagObject = null;

            try
            {
                var bagId = typeof(IPropertyBag).GUID;

                // Get property bag of the moniker
                moniker.BindToStorage(null, null, ref bagId, out bagObject);
                var propertyBag = (IPropertyBag)bagObject;

                // Read FriendlyName
                object value1 = null; object value = null;
                var bresult = propertyBag.Read("DevicePath", ref value1, IntPtr.Zero);
                var hresult = propertyBag.Read("FriendlyName", ref value, IntPtr.Zero);
                if (hresult != 0 )
                {
                    Marshal.ThrowExceptionForHR(hresult);
                }
                return new FullName(value1 as string??string.Empty, value as string ?? string.Empty);
            }
            catch (Exception)
            {
                return new FullName(string.Empty, string.Empty);
            }
            finally
            {
                if (bagObject != null)
                {
                    Marshal.ReleaseComObject(bagObject);
                    bagObject = null;
                }
            }
        }
        
    }
    public class FullName
    {
        private string _name;
        
        private string _deviceInstanceId;

        public FullName(string deviceInstanceId, string name)
        {
            var result = Regex.Match(deviceInstanceId, "#([a-z&0-9]*)#");
            if (result != null)
            {
                deviceInstanceId = result.Groups[1].Value;
            }
            this._name = name;
            this._deviceInstanceId = deviceInstanceId;
        }
        /// <summary>
        /// The deviceInstanceId of the video input device. Can be used to store informations about this cam and restore if connected again 
        /// </summary>
        public string DeviceInstanceId => _deviceInstanceId;
        /// <summary>
        /// The name of the video input device.
        /// </summary>
        public string Name => _name;
        public override string ToString()
        {
            return $"{Name} ({DeviceInstanceId})";
        }
    }
}
