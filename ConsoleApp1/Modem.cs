using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace M5ModemInstaller
{
    public static  class Modem
    {
        const Int64 INVALID_HANDLE_VALUE = -1;
        private static bool FindExistingDevice(string HardwareID, out int lastError)
        {
            const int ERROR_INVALID_DATA = 13;
            const int ERROR_INSUFFICIENT_BUFFER = 122;
            lastError = 0;
            Guid id = Guid.Empty;
            IntPtr hDeviceInfoSet = Win32Wrapper.SetupDiGetClassDevs(ref id, IntPtr.Zero, IntPtr.Zero, (uint)(Win32Wrapper.DIGCF.DIGCF_ALLCLASSES | Win32Wrapper.DIGCF.DIGCF_PRESENT));
            if (hDeviceInfoSet.ToInt64() == INVALID_HANDLE_VALUE)
                return false;
            Win32Wrapper.SP_DEVINFO_DATA DeviceInfoData = new Win32Wrapper.SP_DEVINFO_DATA();
            DeviceInfoData.cbSize = (uint)Marshal.SizeOf(DeviceInfoData);

            uint i = 0;
            uint DataType = 0;

            byte[] buffer = new byte[4];
            IntPtr bufferPtr = Marshal.AllocHGlobal(buffer.Length);

            Marshal.Copy(buffer, 0, bufferPtr, buffer.Length);

            uint requiredSize = 0;

            bool error = false;
            while (Win32Wrapper.SetupDiEnumDeviceInfo(hDeviceInfoSet, i, ref DeviceInfoData))
            {
                i = i + 1;
                while (!Win32Wrapper.SetupDiGetDeviceRegistryProperty(hDeviceInfoSet, ref DeviceInfoData, (uint)Win32Wrapper.SPDRP.SPDRP_HARDWAREID,
                    ref DataType, bufferPtr, (uint)buffer.Length, ref requiredSize))
                {
                    if (Marshal.GetLastWin32Error() == ERROR_INVALID_DATA)
                        break;
                    else if (Marshal.GetLastWin32Error() == ERROR_INSUFFICIENT_BUFFER)
                    {
                        //resize buffer
                        Array.Resize(ref buffer, (int)requiredSize);
                        bufferPtr = Marshal.AllocHGlobal(buffer.Length);

                    }
                    else
                    {
                        error = true;
                        break;
                    }

                }

                if (Marshal.GetLastWin32Error() == ERROR_INVALID_DATA)
                    continue;
                if (error)
                    break;

                if (buffer.Length > 0)
                {
                    string str = Marshal.PtrToStringUni(bufferPtr);  
                    if (str.ToUpper() == HardwareID.ToUpper())
                    {
                        return true;
                    }
                }

            }

            Marshal.FreeHGlobal(bufferPtr);
            lastError = Marshal.GetLastWin32Error();

            Win32Wrapper.SetupDiDestroyDeviceInfoList(hDeviceInfoSet);

            return false;
        }
        public static bool InstallModemDriver(string COMPort)
        {
            Win32Wrapper.SP_DEVINFO_DATA DeviceInfoData = new Win32Wrapper.SP_DEVINFO_DATA();
            DeviceInfoData.cbSize = (uint)Marshal.SizeOf(DeviceInfoData);

           
                Guid classguid = new Guid("4d36e96d-e325-11ce-bfc1-08002be10318");

                //little test
                string infFile = "C:\\WINDOWS\\inf\\M5-MDM-LS.inf";
                string HardwareID = "M5PNPLS";
                int lasterror = 0;
                if (FindExistingDevice(HardwareID, out lasterror))
                {
                //INSTALLFLAG_READONLY = 2
                bool restart = false;
                    if (!Win32Wrapper.UpdateDriverForPlugAndPlayDevices(IntPtr.Zero, HardwareID, infFile, 2, out restart))
                        return false;
                }
                else
                {
                    //ERROR_NO_MORE_ITEMS = 259
                    if (lasterror != (int)Win32Wrapper.WinErrors.ERROR_NO_MORE_ITEMS)
                        return false;
                }

                bool res = Win32Wrapper.SetupDiGetINFClass(infFile, ref classguid, "Modem", 256, IntPtr.Zero);

                IntPtr DeviceInfoSet = Win32Wrapper.SetupDiCreateDeviceInfoList(ref classguid, IntPtr.Zero);
                if (DeviceInfoSet.ToInt64() == INVALID_HANDLE_VALUE)
                {
                    return false;
                }

                if (!Win32Wrapper.SetupDiCreateDeviceInfo(DeviceInfoSet, "Modem", ref classguid, "@mdmhayes.inf,%m2700%;Modem M5", IntPtr.Zero, 0x1, ref DeviceInfoData))
                {
                    int err = Marshal.GetLastWin32Error();
                    Win32Wrapper.SetupDiDestroyDeviceInfoList(DeviceInfoSet);
                    return false;
                }
                if (!Win32Wrapper.SetupDiSetDeviceRegistryProperty(DeviceInfoSet, ref DeviceInfoData, (uint)Win32Wrapper.SPDRP.SPDRP_HARDWAREID, HardwareID,
                    (HardwareID.Length + 2) * Marshal.SystemDefaultCharSize))
                {
                    Win32Wrapper.SetupDiDestroyDeviceInfoList(DeviceInfoSet);
                    return false;
                }
                Win32Wrapper.SP_DRVINFO_DATA drvData = new Win32Wrapper.SP_DRVINFO_DATA();
                drvData.cbSize = (uint)Marshal.SizeOf(drvData);
                uint DIF_REMOVE = 0x00000005;
                bool result = true;
                if (!RegisterModem(DeviceInfoSet, ref DeviceInfoData, COMPort, ref drvData))
                {
                    result = false;
                }
                try
                {
                //INSTALLFLAG_FORCE = 1
                bool restart = false;
                if (!Win32Wrapper.UpdateDriverForPlugAndPlayDevices(IntPtr.Zero, HardwareID, infFile, 1, out restart))                  
                {
                    int err = Marshal.GetLastWin32Error();
                    result = false;
                }
                if (!result)
                {
                    Win32Wrapper.SetupDiCallClassInstaller(DIF_REMOVE, DeviceInfoSet, ref DeviceInfoData);
                }
                Win32Wrapper.SetupDiDestroyDeviceInfoList(DeviceInfoSet);
                return result; ;
                }
                catch
                {
                    Win32Wrapper.SetupDiDestroyDeviceInfoList(DeviceInfoSet);
                    return false;
                }
        }
        public static bool UninstallModem(string COMPort)
        {
            Guid classguid = new Guid("4d36e96d-e325-11ce-bfc1-08002be10318");
            IntPtr hDeviceInfoSet = Win32Wrapper.SetupDiGetClassDevs(ref classguid, IntPtr.Zero, IntPtr.Zero, (uint)Win32Wrapper.DIGCF.DIGCF_PRESENT);
            bool matchFound = false;
            if(hDeviceInfoSet != IntPtr.Zero && hDeviceInfoSet.ToInt64() != INVALID_HANDLE_VALUE)
            {
                Win32Wrapper.SP_DEVINFO_DATA devInfoElem = new Win32Wrapper.SP_DEVINFO_DATA();
                uint index = 0;
                devInfoElem.cbSize = (uint)Marshal.SizeOf(devInfoElem);
                while (Win32Wrapper.SetupDiEnumDeviceInfo(hDeviceInfoSet, index, ref devInfoElem))
                {
                    index = index + 1;
                    IntPtr hKeyDev = Win32Wrapper.SetupDiOpenDevRegKey(hDeviceInfoSet, ref devInfoElem, (uint)Win32Wrapper.DICS_FLAG.DICS_FLAG_GLOBAL,0 , (uint)Win32Wrapper.DIREG.DIREG_DRV, (uint)Win32Wrapper.REGKEYSECURITY.KEY_READ);
                    int test = Marshal.GetLastWin32Error();
                    StringBuilder szDevDesc = new StringBuilder(20, 256);
                    if (hKeyDev.ToInt64() != INVALID_HANDLE_VALUE)
                    {
                        uint pData = 256;
                        uint lpType = 0;
                        int res = Win32Wrapper.RegQueryValueEx(hKeyDev, "AttachedTo", 0, out lpType, szDevDesc, ref pData);
                        if (res == (int)Win32Wrapper.WinErrors.ERROR_SUCCESS)
                        {
                            Win32Wrapper.RegCloseKey(hKeyDev);
                            if (COMPort == szDevDesc.ToString())
                            {
                                Console.WriteLine(String.Format("Found :  {0}", COMPort));
                                matchFound = true;
                                uint DIF_REMOVE = 0x00000005;
                                if (!Win32Wrapper.SetupDiCallClassInstaller(DIF_REMOVE, hDeviceInfoSet, ref devInfoElem))
                                {
                                    Win32Wrapper.SetupDiDestroyDeviceInfoList(hDeviceInfoSet);
                                }
                                break;
                            }

                        }
                    }
                }
                        
                
            }
            return matchFound;
        }
        private static bool RegisterModem(IntPtr hDeviceInfoSet, ref Win32Wrapper.SP_DEVINFO_DATA devInfoElem, string PortName, ref Win32Wrapper.SP_DRVINFO_DATA drvData)
        {
            if (!Win32Wrapper.SetupDiRegisterDeviceInfo(hDeviceInfoSet, ref devInfoElem, 0, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero))
                return false;
            IntPtr hKeyDev = Win32Wrapper.SetupDiOpenDevRegKey(hDeviceInfoSet, ref devInfoElem, (uint)Win32Wrapper.DICS_FLAG.DICS_FLAG_GLOBAL, 0, (uint)Win32Wrapper.DIREG.DIREG_DEV, (uint)Win32Wrapper.REGKEYSECURITY.KEY_READ);

            if (hKeyDev.ToInt64() == INVALID_HANDLE_VALUE && Marshal.GetLastWin32Error() == (int)Win32Wrapper.WinErrors.ERROR_KEY_DOES_NOT_EXIST)
            {
                hKeyDev = Win32Wrapper.SetupDiCreateDevRegKey(hDeviceInfoSet, ref devInfoElem, (int)Win32Wrapper.DICS_FLAG.DICS_FLAG_GLOBAL, 0, (int)Win32Wrapper.DIREG.DIREG_DRV, IntPtr.Zero, IntPtr.Zero);
                if (hKeyDev.ToInt64() == INVALID_HANDLE_VALUE)
                    return false;
            }
            bool check = Win32Wrapper.SetupDiRegisterDeviceInfo(hDeviceInfoSet, ref devInfoElem, 0, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            if (!Win32Wrapper.SetupDiGetSelectedDriver(hDeviceInfoSet, ref devInfoElem, ref drvData))
            {
                int err = Marshal.GetLastWin32Error();
                Console.WriteLine("No driver selected");
            }
            Microsoft.Win32.RegistryValueKind RegType = Microsoft.Win32.RegistryValueKind.String;
            int size = (PortName.Length + 1) * Marshal.SystemDefaultCharSize;
            int ret = Win32Wrapper.RegSetValueEx(hKeyDev, "AttachedTo", 0, (uint)RegType, PortName, (uint)size);

            Win32Wrapper.RegCloseKey(hKeyDev);
            return true;
        }

        /// <summary>
        /// Prevent two modems exisitng for M5 this is to avoid confusion, and minimize errors/
        /// If we always deal with one, we only ever have one phonebook to worry about, one modem to deal with, etc
        /// </summary>
        /// <param name="COMPort"></param>
        public static void UninstallModemsNotAttachedToThisCOMPort(string COMPort)
        {
            RegistryKey regkey = Registry.LocalMachine;
            string keyModem = @"SYSTEM\CurrentControlSet\Control\Class\{4D36E96D-E325-11CE-BFC1-08002BE10318}";
            var subkeys = regkey.OpenSubKey(keyModem).GetSubKeyNames();
            foreach (string key in subkeys)
            {
                bool isAttahced = false;
                string[] sb;
                try
                {
                    sb = regkey.OpenSubKey(keyModem + @"\" + key).GetValueNames();
                }
                catch
                {
                    continue;
                }
                foreach (var item in sb)
                {
                    if (item == "AttachedTo")
                    {
                        isAttahced = true;
                        break;
                    }
                }
                if (isAttahced)
                {
                    string AttachedTo = regkey.OpenSubKey(keyModem + "\\" + key).GetValue("AttachedTo").ToString();
                    if (COMPort.Trim().ToUpper() != AttachedTo.Trim().ToUpper())
                    {
                        UninstallModem(AttachedTo);
                    }
                       
                }

            }
            

        }
        public static bool DoesModemExistOnPort(string COMPort)
        {
            RegistryKey regkey = Registry.LocalMachine;
            string keyModem = @"SYSTEM\CurrentControlSet\Control\Class\{4D36E96D-E325-11CE-BFC1-08002BE10318}";
            var subkeys = regkey.OpenSubKey(keyModem).GetSubKeyNames();
            foreach (string key in subkeys)
            {
                bool isAttahced = false;
                string[] sb;
                try
                {
                    sb = regkey.OpenSubKey(keyModem + @"\" + key).GetValueNames();
                }
                catch
                {
                    continue;
                }
                foreach (var item in sb)
                {
                    if (item == "AttachedTo")
                    {
                        isAttahced = true;
                        break;
                    }
                }
                if (isAttahced)
                {
                    string AttachedTo = regkey.OpenSubKey(keyModem + "\\" + key).GetValue("AttachedTo").ToString();
                    if (COMPort.Trim().ToUpper() == AttachedTo.Trim().ToUpper())
                    {
                        string friendlyName= regkey.OpenSubKey(keyModem + "\\" + key).GetValue("FriendlyName").ToString();//just a test
                        return true;
                    }
                }

            }
            return false;

        }
    }
}
