using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace M5ModemInstaller
{
   
    class Program
    {
        const int SUCCESS = 0;
        const int FAILED = 1;
        const int ALREADYINSTALLED = 2;
   
        static int Main(string[] args)
        {

            USB.DeviceProperties device_prop = new USB.DeviceProperties();
            //StreamWriter writer = File.CreateText(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\M5ModemOutput.txt");
             
           // writer.WriteLine("Getting usb device for 403, 718 pid/vid");
            USB.GetUSBDevice(0x0403, 0x7018, ref device_prop);
            string com_port = device_prop.COMPort;
           // writer.WriteLine("USB device found on " + com_port);
            int returnVal = SUCCESS;

            Console.WriteLine("Uninstalling modems from other ports that are not " + com_port);
          //  writer.WriteLine("Uninstalling modems from other ports that are not " + com_port);
            Modem.UninstallModemsNotAttachedToThisCOMPort(com_port);

            if (!Modem.DoesModemExistOnPort(com_port))
            {
             //   writer.WriteLine("Installing modem on " + com_port);
                Console.WriteLine("Installing modem on " + com_port);
                if (Modem.InstallModemDriver(com_port/*ref writer*/))
                {
                 //   writer.WriteLine("Successfully installed modem driver");
                    returnVal = SUCCESS;
                }
                else { returnVal = FAILED; //writer.WriteLine("failed to install modem");
                }

            }
            else
            {
               // writer.WriteLine("modem already installed from before");
                returnVal = ALREADYINSTALLED;

            }

            PhoneBook.CreatePhonebook();
           // writer.Close();
            return returnVal;
        }
    }
  
}
