using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace M5ModemInstaller
{
    public static class USB
    {
        const Int64 INVALID_HANDLE_VALUE = -1;
        const int BUFFER_SIZE = 1024;
        public struct DeviceProperties
        {
            public string FriendlyName;
            public string DeviceDescription;
            public string DeviceType;
            public string DeviceManufacturer;
            public string DeviceClass;
            public string DeviceLocation;
            public string DevicePath;
            public string DevicePhysicalObjectName;
            public string COMPort;
            public string DeviceInstancePath;
        }
        
        public static bool GetUSBDevice(UInt32 VID, UInt32 PID, ref DeviceProperties DP)
        {
            IntPtr IntPtrBuffer = Marshal.AllocHGlobal(BUFFER_SIZE);
            IntPtr h = IntPtr.Zero;
            Win32Wrapper.WinErrors LastError;
            bool Status = false;

            try
            {
                string DevEnum = "USB";
                string ExpectedDeviceID = "VID_" + VID.ToString("X4") + "&" + "PID_" + PID.ToString("X4");
                ExpectedDeviceID = ExpectedDeviceID.ToLowerInvariant();

                h = Win32Wrapper.SetupDiGetClassDevs(IntPtr.Zero, DevEnum, IntPtr.Zero, (int)(Win32Wrapper.DIGCF.DIGCF_PRESENT | Win32Wrapper.DIGCF.DIGCF_ALLCLASSES));
                if (h.ToInt64() != INVALID_HANDLE_VALUE)
                {
                    bool Success = true;
                    uint i = 0;
                    while (Success)
                    {
                        if (Success)
                        {
                            UInt32 RequiredSize = 0;
                            UInt32 RegType = 0;
                            IntPtr Ptr = IntPtr.Zero;

                            //Create a Device Info Data structure
                            Win32Wrapper.SP_DEVINFO_DATA DevInfoData = new Win32Wrapper.SP_DEVINFO_DATA();
                            DevInfoData.cbSize = (uint)Marshal.SizeOf(DevInfoData);
                            Success = Win32Wrapper.SetupDiEnumDeviceInfo(h, i, ref DevInfoData);

                            if (Success)
                            {
                                //Get the required buffer size
                                //First query for the size of the hardware ID, so we can know how big a buffer to allocate for the data.
                                Win32Wrapper.SetupDiGetDeviceRegistryProperty(h, ref DevInfoData, (UInt32)Win32Wrapper.SPDRP.SPDRP_HARDWAREID, ref RegType, IntPtr.Zero, 0, ref RequiredSize);

                                LastError = (Win32Wrapper.WinErrors)Marshal.GetLastWin32Error();
                                if (LastError == Win32Wrapper.WinErrors.ERROR_INSUFFICIENT_BUFFER)
                                {
                                    if (RequiredSize > BUFFER_SIZE)
                                    {
                                        Status = false;
                                    }
                                    else
                                    {
                                        if (Win32Wrapper.SetupDiGetDeviceRegistryProperty(h, ref DevInfoData, (UInt32)Win32Wrapper.SPDRP.SPDRP_HARDWAREID, ref RegType, IntPtrBuffer, BUFFER_SIZE, ref RequiredSize))
                                        {
                                            string HardwareID = Marshal.PtrToStringAuto(IntPtrBuffer);
                                            HardwareID = HardwareID.ToLowerInvariant();
                                            if (HardwareID.Contains(ExpectedDeviceID))
                                            {
                                                Status = true; //Found device
                                                if (Win32Wrapper.SetupDiGetDeviceRegistryProperty(h, ref DevInfoData, (UInt32)Win32Wrapper.SPDRP.SPDRP_FRIENDLYNAME, ref RegType, IntPtrBuffer, BUFFER_SIZE, ref RequiredSize))
                                                {
                                                    DP.FriendlyName = Marshal.PtrToStringAuto(IntPtrBuffer);
                                                }
                                                if (Win32Wrapper.SetupDiGetDeviceRegistryProperty(h, ref DevInfoData, (UInt32)Win32Wrapper.SPDRP.SPDRP_DEVTYPE, ref RegType, IntPtrBuffer, BUFFER_SIZE, ref RequiredSize))
                                                {
                                                    DP.DeviceType = Marshal.PtrToStringAuto(IntPtrBuffer);
                                                }
                                                if (Win32Wrapper.SetupDiGetDeviceRegistryProperty(h, ref DevInfoData, (UInt32)Win32Wrapper.SPDRP.SPDRP_CLASS, ref RegType, IntPtrBuffer, BUFFER_SIZE, ref RequiredSize))
                                                {
                                                    DP.DeviceClass = Marshal.PtrToStringAuto(IntPtrBuffer);
                                                }
                                                if (Win32Wrapper.SetupDiGetDeviceRegistryProperty(h, ref DevInfoData, (UInt32)Win32Wrapper.SPDRP.SPDRP_MFG, ref RegType, IntPtrBuffer, BUFFER_SIZE, ref RequiredSize))
                                                {
                                                    DP.DeviceManufacturer = Marshal.PtrToStringAuto(IntPtrBuffer);
                                                }
                                                if (Win32Wrapper.SetupDiGetDeviceRegistryProperty(h, ref DevInfoData, (UInt32)Win32Wrapper.SPDRP.SPDRP_LOCATION_INFORMATION, ref RegType, IntPtrBuffer, BUFFER_SIZE, ref RequiredSize))
                                                {
                                                    DP.DeviceLocation = Marshal.PtrToStringAuto(IntPtrBuffer);
                                                }
                                                if (Win32Wrapper.SetupDiGetDeviceRegistryProperty(h, ref DevInfoData, (UInt32)Win32Wrapper.SPDRP.SPDRP_LOCATION_PATHS, ref RegType, IntPtrBuffer, BUFFER_SIZE, ref RequiredSize))
                                                {
                                                    DP.DevicePath = Marshal.PtrToStringAuto(IntPtrBuffer);
                                                }
                                                if (Win32Wrapper.SetupDiGetDeviceRegistryProperty(h, ref DevInfoData, (UInt32)Win32Wrapper.SPDRP.SPDRP_PHYSICAL_DEVICE_OBJECT_NAME, ref RegType, IntPtrBuffer, BUFFER_SIZE, ref RequiredSize))
                                                {
                                                    DP.DevicePhysicalObjectName = Marshal.PtrToStringAuto(IntPtrBuffer);
                                                }
                                                if (Win32Wrapper.SetupDiGetDeviceRegistryProperty(h, ref DevInfoData, (UInt32)Win32Wrapper.SPDRP.SPDRP_DEVICEDESC, ref RegType, IntPtrBuffer, BUFFER_SIZE, ref RequiredSize))
                                                {
                                                    DP.DeviceDescription = Marshal.PtrToStringAuto(IntPtrBuffer);
                                                }
                                                StringBuilder sb = new StringBuilder(BUFFER_SIZE);
                                                if (Win32Wrapper.SetupDiGetDeviceInstanceId(h, ref DevInfoData, sb, BUFFER_SIZE, ref RequiredSize))
                                                {
                                                    DP.DeviceInstancePath = sb.ToString();
                                                }
                                                string path = DP.DeviceInstancePath;
                                                string key = path.Replace("USB\\", "");
                                                key = key.Replace("\\", "+");
                                                key = key.Remove(0, 17);
                                                string deviceID = "VID_" + VID.ToString("X4") + "+" + "PID_" + PID.ToString("X4");
                                                key = deviceID + key;

                                                DP.COMPort = (string)Registry.GetValue(@"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\FTDIBUS\" + key + @"\0000\Device Parameters",
                                                    "PortName", null);                                          
                                                
                                                break;
                                            }
                                            else
                                            {
                                                Status = false;
                                            } //End of if (HardwareID.Contains(ExpectedDeviceID))
                                        }
                                    } //End of if (RequiredSize > BUFFER_SIZE)
                                } //End of if (LastError == Win32Wrapper.WinErrors.ERROR_INSUFFICIENT_BUFFER)
                            } // End of if (Success)
                        } // End of if (Success)
                        else
                        {
                            LastError = (Win32Wrapper.WinErrors)Marshal.GetLastWin32Error();
                            Status = false;
                        }
                        i++;
                    } // End of while (Success)
                } //End of if (h.ToInt64() != INVALID_HANDLE_VALUE)

                return Status;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Win32Wrapper.SetupDiDestroyDeviceInfoList(h); //Clean up the old structure we no longer need.
                Marshal.FreeHGlobal(IntPtrBuffer);
            }
        }
    }
}
