using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DotRas;
namespace M5ModemInstaller
{
    public static class PhoneBook
    {
        public static void CreatePhonebook()
        {
            string programData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            string M5Folder = programData + "\\M5";
            if (!Directory.Exists(M5Folder))
                Directory.CreateDirectory(M5Folder);

            string phoneBookPath = M5Folder + "\\M5.pbk";

       
            CreateRas(phoneBookPath);
        }
        
        //This creates the M5 phonebook with all the necessary parameters in it
        private static void CreateRas(string phonebookPath)
        {
            DotRas.RasPhoneBook pb = new DotRas.RasPhoneBook();
            pb.Open(phonebookPath);
            RasDevice rasDevice = RasDevice.GetDevices().Where(P => P.DeviceType == RasDeviceType.Modem && P.Name.ToUpper().StartsWith("M5-")).FirstOrDefault();
            if (rasDevice == null)
                return;
           
            RasEntry entry = pb.Entries.Where(p => p.Name == "M5_1").FirstOrDefault();
            if (entry == null)
                entry = RasEntry.CreateDialUpEntry("M5_1", "0", rasDevice);

            entry.FramingProtocol = RasFramingProtocol.Ppp;
            entry.EncryptionType = RasEncryptionType.None;
            entry.FrameSize = 0;

            entry.Options.DisableLcpExtensions = true;
            entry.Options.DisableNbtOverIP = true;
            entry.Options.DoNotNegotiateMultilink = true;
            entry.Options.DoNotUseRasCredentials = true;
            entry.Options.Internet = false;
            entry.Options.IPHeaderCompression = false;
            entry.Options.ModemLights = true;
            entry.Options.NetworkLogOn = false;

            entry.Options.PreviewDomain = false;
            entry.Options.PreviewPhoneNumber = false;
            entry.Options.PreviewUserPassword = false;
            entry.Options.PromoteAlternates = false;

            entry.Options.ReconnectIfDropped = false;
            entry.Options.RemoteDefaultGateway = false;

            entry.Options.RequireChap = false;
            entry.Options.RequireDataEncryption = false;
            entry.Options.RequireEap = false;
            entry.Options.RequireEncryptedPassword = false;
            entry.Options.RequireMSChap = false;
            entry.Options.RequireMSChap2 = false;
            entry.Options.RequireMSEncryptedPassword = false;
            entry.Options.RequirePap = false;
            entry.Options.RequireSpap = false;
            entry.Options.RequireWin95MSChap = false;

            entry.Options.SecureClientForMSNet = true;
            entry.Options.SecureFileAndPrint = true;
            entry.Options.SecureLocalFiles = true;

            entry.Options.SharedPhoneNumbers = false;
            entry.Options.SharePhoneNumbers = false;
            entry.Options.ShowDialingProgress = false;

            entry.Options.SoftwareCompression = false;
            entry.Options.TerminalAfterDial = false;
            entry.Options.TerminalBeforeDial = false;

            entry.Options.UseCountryAndAreaCodes = false;
            entry.Options.UseGlobalDeviceSettings = false;
            entry.Options.UseLogOnCredentials = false;
            entry.Options.UsePreSharedKey = false;
    
            if(pb.Entries.Contains(entry.Name))
            {
                entry.Update();
            }
            else
            {
                pb.Entries.Add(entry);
            }

        }
    }
}
