using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Pulse;

namespace SendToMachine
{
    class Program
    {
        private static Random random = new Random();

        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
                .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        static void Main(string[] args)
        {
            //var randomString = RandomString(6);

            //IApplication app = new Pulse.Application();
            //try
            //{
            //    Console.WriteLine("SDK Version: " + app.Version);
            //    IEmbDesign myDesign = app.NewDesign("Normal", "Barudan Z");
            //    try
            //    {
            //        // Create a line of embroidery text in the Block New font.  The height = 1 inch (254 embroidery points, 1 embroidery point = .1 mm)
            //        TextProperties prop = app.NewTextProperties();
            //        Console.WriteLine("Creating Lettering");
            //        myDesign.AddLineText(randomString, "Block New", 254, 0, 0, 0, 0, JustifyTypes.jtCenter, EnvelopeTypes.etRectangle, 0, prop);

            //        // Create a bitmap image of the design for preview
            //        IBitmapImage myImage = app.NewImage(300, 300);
            //        try
            //        {
            //            Console.WriteLine("Saving Image");
            //            myDesign.Render(myImage, 0, 0, 300, 300);
            //            myImage.Save(@"C:\Users\chrisf\Documents\Pulse Micro\Tajima\Designs\PulseEmbCom\EmbImage.png", ImageTypes.itAuto);

            //            // Connect to an embroidery design spooler on the local machine.  The design spooler controls embroidery machine connections.
            //            IDesignSpooler spooler = null;
            //            try
            //            {
            //                spooler = app.ConnectToDesignSpooler("localhost", 9050);

            //                Console.WriteLine("Gettling Machine List");

            //                // Get a list of embroidery machines connected to the spooler
            //                IMachines MachineList = spooler.GetMachines();
            //                try
            //                {
            //                    for (int i = 0; i <= MachineList.Count - 1; i++)
            //                    {
            //                        IMachine machine = MachineList[i];
            //                        try
            //                        {
            //                            if (machine != null)
            //                            {
            //                                Console.WriteLine(machine.Name);
            //                                // Remove any existing designs from the queue for this machine

            //                                Console.WriteLine("Sending Design");

            //                                machine.EmptyQueue();

            //                                // Send the the Emb Design to the design queue with a download count of 1

            //                                machine.PublishDesign(myDesign, randomString, 1);
            //                                // Sending the lettering to the fist machine it finds on the list of machines
            //                                Console.WriteLine("Design Sent to machine " + machine.Name);
            //                            }
            //                        }
            //                        finally
            //                        {
            //                            System.Runtime.InteropServices.Marshal.ReleaseComObject(machine);
            //                        }
            //                        // Send the design to the first machine in the list
            //                    }
            //                }
            //                finally
            //                {
            //                    System.Runtime.InteropServices.Marshal.ReleaseComObject(MachineList);
            //                }
            //            }
            //            catch (Exception e)
            //            {
            //                Console.WriteLine("Unable to connect to design spooler.  Is design spooler running?");
            //                Console.WriteLine(e.Message);
            //            }
            //            finally
            //            {
            //                if (spooler != null)
            //                    System.Runtime.InteropServices.Marshal.ReleaseComObject(spooler);
            //            }
            //        }
            //        finally
            //        {
            //            System.Runtime.InteropServices.Marshal.ReleaseComObject(myImage);
            //        }
            //    }
            //    finally
            //    {
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(myDesign);
            //    }
            //}
            //finally
            //{
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            //}

            if (args.Length <2)
                Console.WriteLine("There must be at least 2 arguments. The machine name and the machine format should be supplied.  ConsoleApp5.exe BEKT-S1501CAII TAJIMA");
            else
                SendToSerialMachine(args[0], args[1]);

            Console.WriteLine();

            Console.ReadLine();
        }

        private static void SendToSerialMachine(string machineName, string machineformat)
        {
            var randomString = RandomString(6);

            int numStitches;
            Pulse.IApplication app = new Pulse.Application();
            Pulse.IEmbDesign design = null;
            Pulse.IMachine machine = null;
            IEmbDesign myDesign = app.NewDesign("Normal", machineformat);
            TextProperties prop = app.NewTextProperties();
            try
            {
                Console.WriteLine("Creating Lettering");
                myDesign.AddLineText(randomString, "Block New", 254, 0, 0, 0, 0, JustifyTypes.jtCenter, EnvelopeTypes.etRectangle, 0, prop);
                var localPath = AppDomain.CurrentDomain.BaseDirectory;
                var embroideryFile = $"{localPath}{randomString}.pcf";
                myDesign.Save(embroideryFile, FileTypes.ftAuto);

                design = LoadDesign(app, embroideryFile);
                if (design == null) return;

                machine = GetMachine(app, machineName);
                if (machine == null) return;

                numStitches = Convert.ToInt32(design.NumStitches);
                ClearDesignsQueue(machine);

                // Colourize the design
                //if (param.SetNeedles)
                //    SetNeedles(design, param, db, log);

                //var jobToSend = TruncateJob(param.Job, param.TruncateLength);

                try
                {
                    Console.WriteLine($"Publishing design with a name '{randomString}' to the machine '{machineName}'");
                    machine.PublishDesign(design, randomString, 1);
                    Console.WriteLine("Design sent successfully");
                }
                catch (Exception e1)
                {
                    Console.WriteLine($"Error sending design to machine {machineName}: {e1.InnerException?.Message ?? e1.Message}");
                    return;
                }

                Console.WriteLine($"Deleting file {embroideryFile}");
                if (System.IO.File.Exists(embroideryFile))
                    System.IO.File.Delete(embroideryFile);
            }
            finally
            {
                if (machine != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(machine);

                if (design != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(design);

                if (prop != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(prop);

                if (myDesign != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(myDesign);

                if (app != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            }
        }

        private static Pulse.IEmbDesign LoadDesign(Pulse.IApplication app, string embroideryFile)
        {
            Pulse.IEmbDesign design = null;
            try
            {
                Console.WriteLine($"Loading design from {embroideryFile}...");
                design = app.OpenDesign(embroideryFile, Pulse.FileTypes.ftPXF, Pulse.OpenTypes.otOutlines, "Tajima");

                return design;
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error loading design: {e.InnerException?.Message ?? e.Message}");
                return null;
            }
        }

        private static Pulse.IMachine GetMachine(Pulse.IApplication app, string machineName)
        {
            Pulse.IDesignSpooler spooler = null;
            try
            {
                Console.WriteLine("Connecting to Design Spooler...");
                spooler = app.ConnectToDesignSpooler("localhost", 9050);

                Pulse.IMachines machines = null;
                try
                {
                    Console.WriteLine("Getting machines from Design Spooler...");
                    machines = spooler.GetMachines();

                    if (machines != null && machines.Count > 0)
                        Console.WriteLine($"Got this many machines from Design Spooler: {machines.Count}");
                    else
                        Console.WriteLine("There were no machines found in Design Spooler");

                    Console.WriteLine($"Getting machine '{machineName}' from Design Spooler...");
                    var machine = machines[machineName];

                    return machine;
                }
                catch (Exception e1)
                {
                    Console.WriteLine($"Error getting machine '{machineName}' from Design Spooler: {e1.InnerException?.Message ?? e1.Message}");
                    return null;
                }
                finally
                {
                    if (machines != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(machines);
                }
            }
            catch (Exception e2)
            {
                Console.WriteLine($"Error while establishing connection to Design Spooler: {e2.InnerException?.Message ?? e2.Message}");
                return null;
            }
            finally
            {
                if (spooler != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(spooler);
            }
        }

        private static void ClearDesignsQueue(Pulse.IMachine machine)
        {
            Console.WriteLine("Getting the queued designs from the machine");
            var queuedDesigns = machine.GetQueuedDesigns();
            try
            {
                for (int i = 0; i < queuedDesigns.Count; i++)
                {
                    Console.WriteLine($"Clearing the queued designs from the machine at index {i}");
                    queuedDesigns[i].Delete();
                }
            }
            finally
            {
                if (queuedDesigns != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(queuedDesigns);
            }
        }
    }
}
