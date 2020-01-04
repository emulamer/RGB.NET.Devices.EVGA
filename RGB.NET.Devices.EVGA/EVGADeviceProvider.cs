using Microsoft.Win32;
using RGB.NET.Core;
using RGB.NET.Devices.EVGA.Generic;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using System.IO.Pipes;

namespace RGB.NET.Devices.EVGA
{
    public class EVGADeviceProvider : IRGBDeviceProvider
    {
        private string _evgaPath = null;
        private string _precisionExe = null;
        private string _precisionExeName = null;
        private Type _comInterface;
        private Process _proxyProc;
        private NamedPipeClientStream nps;
        private static byte[] MAGIC = new byte[] { 0xB, 0xE, 0xE, 0xF };
        private byte[] ReadMsg(Stream nps)
        {
            byte[] magic = new byte[4];
            int readlen = nps.Read(magic, 0, 4);
            if (readlen != 4 || !(magic[0] == 0xB && magic[1] == 0xE && magic[2] == 0xE && magic[3] == 0xF))
            {
                throw new Exception("No beef");
            }

            byte[] bMsg = new byte[16];
            readlen = nps.Read(bMsg, 0, bMsg.Length);
            if (readlen != 16)
            {
                throw new Exception("Bad data");
            }

            return bMsg;
        }
        private int GetDeviceCount()
        {
            nps.Write(MAGIC, 0, 4);
            byte[] bMsg = new byte[16];
            bMsg[0] = 1;
            nps.Write(bMsg, 0, bMsg.Length);
            bMsg = ReadMsg(nps);
            if (bMsg[0] != 2)
            {
                throw new Exception("Wrong response");
            }
            return bMsg[1];
        }

        private int GetLedCount(int deviceId)
        {
            nps.Write(MAGIC, 0, 4);
            var bMsg = new byte[16];
            bMsg[0] = 5;
            bMsg[1] = (byte)deviceId;
            nps.Write(bMsg, 0, bMsg.Length);
            bMsg = ReadMsg(nps);
            if (bMsg[0] != 6)
            {
                throw new Exception("Wrong response");
            }
            return bMsg[1];
        }
        public static void Log(string msg)
        {
            try
            {
                File.AppendAllText(Path.Combine(Path.GetTempPath(), "evgaproxy.log"), DateTime.Now.ToString() + ": " + (msg ?? "") + "\r\n");
            }
            catch
            {
                System.Threading.Thread.Sleep(100);
                try
                {
                    File.AppendAllText(Path.Combine(Path.GetTempPath(), "evgaproxy.log"), DateTime.Now.ToString() + ": " + (msg ?? "") + "\r\n");
                }
                catch { }
            }
        }
        internal void SetLed(int deviceId, int ledId, byte a, byte r, byte g, byte b)
        {
            var bMsg = new byte[16];
            bMsg[0] = 10;
            bMsg[1] = (byte)deviceId;
            bMsg[2] = (byte)ledId;
            bMsg[3] = a;
            bMsg[4] = r;
            bMsg[5] = g;
            bMsg[6] = b;
            nps.Write(MAGIC, 0, 4);
            nps.Write(bMsg, 0, bMsg.Length);
        }

        public EVGADeviceProvider()
        {
            Log("Constructor hit");
            try
            {
                Log("Trying to make temp folder");
                var tempFolder = Path.Combine(Path.GetTempPath(), "evgaproxy");
                if (!Directory.Exists(tempFolder))
                {
                    Directory.CreateDirectory(tempFolder);
                }
                try
                {
                    string proxyPath = Path.Combine(tempFolder, "EVGAProxy.exe");
                    using (var s = Assembly.GetExecutingAssembly().GetManifestResourceStream("RGB.NET.Devices.EVGA.EVGAProxy.exe"))
                    {
                        using (var fs = File.Open(proxyPath, FileMode.Create))
                        {
                            s.CopyTo(fs);
                        }
                    }
                    ProcessStartInfo psi = new ProcessStartInfo(proxyPath) { UseShellExecute = true, WorkingDirectory = tempFolder, CreateNoWindow = true };
                    _proxyProc = Process.Start(psi);
                }
                catch (Exception ex)
                {
                    //couldn't write it, maybe it's already there and running?
                }
                nps = new NamedPipeClientStream(".", "evgargbled", PipeDirection.InOut);
                try
                {
                    nps.Connect(5000);
                }
                catch
                {
                    throw new Exception("Coudln't connect to EVGAProxy");
                }
                EVGADeviceInfo devInfo = new EVGADeviceInfo();
                _devices.Clear();

                var numDevs = GetDeviceCount();
                for (int i = 0; i < numDevs; i++)
                {
                    int ledCount = GetLedCount(i);
                    _devices.Add(new EVGADevice(devInfo, SetLed, (uint)i, (uint)ledCount));
                }
            }
            catch (Exception ex)
            {
                Log($"Exception in rgb.net: {ex.Message}");
                throw;
            }
        }

        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            var name = args.Name.Split(',')[0];
            var assName = Path.Combine(_evgaPath, name + ".dll");
            if (!File.Exists(assName))
            {
                assName = Path.Combine(_evgaPath, name + ".exe");
            }
            if (File.Exists(assName))
            {
                try
                {
                    var a = Assembly.Load(File.ReadAllBytes(assName));
                    return a;
                }
                catch (Exception ex)
                {
                }
            }
            return LoadFromEvgaResource(name);
        }

        private Assembly evgaAssembly;
        private List<string> evgaResources = null;
        private Assembly LoadFromEvgaResource(string name)
        {
            try
            {
                if (evgaAssembly == null)
                {
                    evgaAssembly = Assembly.ReflectionOnlyLoadFrom(_precisionExe);
                }
                if (evgaResources == null)
                {
                    evgaResources = new List<string>();
                    evgaResources.AddRange(evgaAssembly.GetManifestResourceNames());
                }
                var found = evgaResources.FirstOrDefault(x => x.ToLower().EndsWith(name.ToLower() + ".dll"));
                if (found == null)
                {
                    return null;
                }
                using (var mrs = evgaAssembly.GetManifestResourceStream(found))
                {
                    using (var ms = new MemoryStream())
                    {
                        mrs.CopyTo(ms);
                        return Assembly.Load(ms.ToArray());
                    }
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private static EVGADeviceProvider _instance;
        
        public static EVGADeviceProvider Instance => _instance ?? new EVGADeviceProvider();

        public bool IsInitialized => true;

        private List<IRGBDevice> _devices = new List<IRGBDevice>();
        public IEnumerable<IRGBDevice> Devices => _devices;

        public bool HasExclusiveAccess => false;

        public void Dispose()
        {
            if (_proxyProc != null)
            {
                try
                {
                    _proxyProc.Kill();
                } catch
                { }
                finally
                {
                    _proxyProc = null;
                }
            }
        }

        public bool Initialize(RGBDeviceType loadFilter = (RGBDeviceType)(-1), bool exclusiveAccessIfPossible = false, bool throwExceptions = false)
        {
            return true;
        }

        public void ResetDevices()
        {
            
        }
    }
}
