using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


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
            USB.GetUSBDevice(0x0403, 0x7018, ref device_prop);
            string com_port = device_prop.COMPort;
            int returnVal = SUCCESS;

            Console.WriteLine("Uninstalling modems from other ports that are not " + com_port);
            
            Modem.UninstallModemsNotAttachedToThisCOMPort(com_port);
       
            if (!Modem.DoesModemExistOnPort(com_port))
            {
                Console.WriteLine("Installing modem on " + com_port);
                if (Modem.InstallModemDriver(com_port))
                {

                    returnVal = SUCCESS;
                }
                else { returnVal = FAILED; }
             
            }
            else
                returnVal =  ALREADYINSTALLED;

            PhoneBook.CreatePhonebook();
    
            return returnVal;
        }
    }
  
}
