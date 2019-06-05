# .NET-Troubleshoot-AddIn
The “Troubleshoot Add-In Sample Code” was created to automatically capture information, such as steps-to-reproduce, and workstation information in one fell swoop instead of requiring agents to install or use different tools outside of Oracle Service Cloud. All that agents need to do is push the "start" button (located in the status bar) and push the “stop” when the agents have completed all steps to reproduce the error.

This sample code is started by receiving ServerConfigProperty from the TroubleshootStatusBarAddIn.cs followed by two methods (1) to define where the results will be saved and (2) to start the windows standard application to capture steps to reproduce.

 

public StatusBarControl(IGlobalContext globalContext, bool isOscInfo, bool isScreenCap, 
        string tspath, int reminderInMinutes, string initialNotification, string finalNotificaiton)
{
        InitializeComponent();
        _osvcLoggedin = globalContext.Login;
        _osvcInterface = globalContext.InterfaceName;
        _osvcSitename = globalContext.InterfaceURL;
        GContext = globalContext;

        _isOscInfo = isOscInfo;
        _isScreenCap = isScreenCap;
        _tspath = tspath;
        _reminderInMinutes = reminderInMinutes;
        _nextConfirmation = _reminderInMinutes;
        _initialnotification = initialNotification;
        _finalnotification = finalNotificaiton;
}

private void btnStart_Click(object sender, EventArgs e)
{
        try
        { 
                btnStart.Enabled = false;
                btnStop.Enabled = true;

                var sitename = new Uri(_osvcSitename);
                _hostname = sitename.Host;

                TroubleshootDirectory();

                using (var frmWaitForm = new FrmWaitForm(LoadTroubleshoot, _initialnotification))
                {
                        frmWaitForm.ShowDialog();
                }

                _isActive = true;
        }
        catch (Exception ex)
        {                
                MessageBox.Show(ex.Message);
        }
}

private static void TroubleshootDirectory()
{
        try
        {
                if (_tspath == null)
                {
                        Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\OSvC_Troubleshoot\\" + _osvcLoggedin + "\\" + _osvcLoggedin + "_" + DateTime.Now.ToString("MMddyyyyHHmmss"));
                        _path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\OSvC_Troubleshoot\\" +  _osvcLoggedin + "\\" + _osvcLoggedin + "_" + DateTime.Now.ToString("MMddyyyyHHmmss");
                        _rootpath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\OSvC_Troubleshoot\\";
                }
                else
                {
                        if (Directory.Exists(_tspath))
                        {
                                Directory.CreateDirectory(_tspath + "\\OSvC_Troubleshoot\\" + "\\" + _osvcLoggedin + "\\" + _osvcLoggedin + "_" + DateTime.Now.ToString("MMddyyyyHHmmss"));
                                _path = _tspath + "\\OSvC_Troubleshoot\\" + _osvcLoggedin + "\\" + _osvcLoggedin + "_" + DateTime.Now.ToString("MMddyyyyHHmmss");
                                _rootpath = _tspath + "\\OSvC_Troubleshoot\\";
                        }
                        else
                        {
                                Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\OSvC_Troubleshoot\\" + _osvcLoggedin + "\\" + _osvcLoggedin + "_" + DateTime.Now.ToString("MMddyyyyHHmmss"));
                                _path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\OSvC_Troubleshoot\\" + _osvcLoggedin + "\\" + _osvcLoggedin + "_" + DateTime.Now.ToString("MMddyyyyHHmmss");
                                _rootpath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\OSvC_Troubleshoot\\";
                        }
                }
        }
        catch (Exception e)
        {                
                MessageBox.Show(e.Message);
        }
}

 

The sample code is centralizing all actions to start the external application in a single method. As now, it is loading only PSR, but you can add Fiddler as described before and also make it an option trought ServerConfigProperty.

 

private void LoadTroubleshoot()
{
        try
        {
                if (_isScreenCap)
                        StartPsr();

                for (var i = 0; i <= 500; i++)
                        Thread.Sleep(10);
        }
        catch (Exception e)
        {
                GContext.LogMessage(e.Message);
        }
}
 

The following method is called to start Windows PSR.

 

private void StartPsr()
{
        try
        {
                TryKillProcess(PsrExe);

                _psrLogFile = Path.Combine(_path + "\\" + DateTime.Now.ToString("MMddyyyyHHmmss") + "_Steps_To_Reproduce.zip");
                InvokeProcess(PsrExe,
                        $"/start /gui 0 /output \"{_psrLogFile}\" /slides 1 /recordpid \"{Process.GetCurrentProcess().Id}\"");

        }
        catch (Exception e)
        {
                GContext.LogMessage(e.Message);
        }
}
 

The following two methods are generic, so you can reuse that for other applications if it is needed. These methods are used to start and stop Windows PSR in this context.

private static void TryKillProcess(string processName)
{
        var processes = Process.GetProcessesByName(processName);
        foreach (var proc in processes)
        {
                try
                {
                        proc.Kill();
                }
                catch (Exception e)
                {                    
                        MessageBox.Show(e.Message);
                }
        }
}

private static Process InvokeProcess(string processName, string parameters)
{
        var startInfo = new ProcessStartInfo(processName, parameters)
        {
                WindowStyle = ProcessWindowStyle.Hidden,
                UseShellExecute = true,
                ErrorDialog = false
        };

        var proc = new Process {StartInfo = startInfo};
        proc.Start();

        return proc;
}
 

The following method will take care of the time control, plus will control the duration. It will remind the agent that the capture is still running in case the agent has accidentally forgotten to stop.

private void timer1_Tick_1(object sender, EventArgs e)
{
        if (_isActive)
        {
                _seconds++;

                if (_seconds > 59)
                {
                        _minutes++;
                        _seconds = 0;
                        if (Convert.ToInt32(_minutes) >= _nextConfirmation)
                        {
                                _nextConfirmation += _reminderInMinutes;

                                var dialogResult = MessageBox.Show(Resources.StatusBarControl_timer1_Tick_1_Are_you_still_collecting_data_, Resources.StatusBarControl_timer1_Tick_1_Troubleshoot, MessageBoxButtons.YesNo);
                                switch (dialogResult)
                                {
                                        case DialogResult.Yes:
                                                return;
                                        case DialogResult.No:
                                                btnStop_Click(sender, e);
                                                break;
                                        case DialogResult.None:
                                                btnStop_Click(sender, e);
                                                break;
                                        case DialogResult.OK:
                                                btnStop_Click(sender, e);
                                                break;
                                        case DialogResult.Cancel:
                                                btnStop_Click(sender, e);
                                                break;
                                        case DialogResult.Abort:
                                                btnStop_Click(sender, e);
                                                break;
                                        case DialogResult.Retry:
                                                btnStop_Click(sender, e);
                                                break;
                                        case DialogResult.Ignore:
                                                btnStop_Click(sender, e);
                                                break;
                                        default:
                                                throw new ArgumentOutOfRangeException();
                                }
                        }

                }
                if (_minutes >= 59)
                {
                        _hours++;
                        _minutes = 0;
                }
        }

        lblTimer.Text = string.Format(AppendZero(_hours) + ":" + AppendZero(_minutes) + ":" + AppendZero(_seconds));
}

private static string AppendZero(double str)
{
        if (str <= 9)
                return "0" + str;
        return str.ToString(CultureInfo.CurrentCulture);
}
 

When the Stop button is clicked the following method is called.

private void btnStop_Click(object sender, EventArgs e)
{
        try
        {
                btnStart.Enabled = true;
                btnStop.Enabled = false;

                using (var frmWaitForm = new FrmWaitForm(StopTroubleshoot, _initialnotification))
                {
                        frmWaitForm.ShowDialog();
                }
        }
        catch (Exception ex)
        {
                GContext.LogMessage(ex.Message);
        }            
}      
 

Similar to the logic applied to the start load, here the sequence to stop the applications running will take place.

private void StopTroubleshoot()
{
        try
        {
                btnStop.Enabled = false;
                btnStart.Enabled = true;
                _isActive = false;
                _seconds = _minutes = _hours = 0;

                if (_isScreenCap)
                        StopPsr();

                if (_isOscInfo)
                {                    
                        Task.Factory.StartNew(() =>
                        {
                                var ocsFile = _path + "\\OSvCinfo.bat"; // File downloaded in 03/20/2018.
                                if (!File.Exists(ocsFile))                                                    
                                        ExtractEmbeddedResource("Troubleshoot_StatusBar", _rootpath, "OSvCInfo", "OSvCinfo.bat");                        
                        }).ContinueWith(antecedent =>
                        {                        
                                RunOsvCInfo();
                        });
                }
                else
                {
                        Task.Factory.StartNew(WorkStationInfo).ContinueWith(antecedent =>
                        {
                                CompressAndNotify();
                        });                    
                }
        }
        catch (Exception e)
        {
                GContext.LogMessage(e.Message);
        }
}
 
The generic kill and invoke method will be called again to stop Windows PSR.

private void StopPsr()
{
        try
        {

                InvokeProcess(PsrExe, @"/stop").WaitForExit(60000);

                // Sleep to ensure PSR completes file creation operation
                Thread.Sleep(SaveDelay);

                if (!File.Exists(_psrLogFile))
                {
                        MessageBox.Show(Resources.StatusBarControl_LoadPsr_No_user_actions_were_recorded_by_PSR_);
                }
        }
        catch (Exception e)
        {
                GContext.LogMessage(e.Message);
        }
}
 

If the ServerConfigProperty that allows OSCInfo.bat file runs has not been defined as True, this piece of code will collect at least the basic workstation information.

 

private static void WorkStationInfo()
{
        try
        {
                var compInfor = new Microsoft.VisualBasic.Devices.ComputerInfo();

                var ocsInfo = _path + "\\" + DateTime.Now.ToString("MMddyyyyHHmmss") + "_OSvCInfo.txt";                

                using (var sw = File.CreateText(ocsInfo))
                {
                        var openSubKey = Registry.LocalMachine.OpenSubKey("hardware\\description\\system\\centralprocessor\\0");                   

                        sw.WriteLine("-- Start Workstation Information --");
                        sw.WriteLine("OSvC Site Name: " + _osvcInterface);
                        sw.WriteLine("Host Name: " + Environment.MachineName);
                        if (openSubKey != null)
                                sw.WriteLine("Processor: " + openSubKey.GetValue("ProcessorNameString"));
                        sw.WriteLine("OS: " + compInfor.OSFullName);
                        sw.WriteLine("Available Physical Memory: " + compInfor.AvailablePhysicalMemory / 1024 / 1024 / 1024 + "GB");
                        sw.WriteLine("Available Virtual Memory: " + compInfor.AvailableVirtualMemory / 1024 / 1024 / 1024 + "GB");
                        sw.WriteLine("Total Physical Memory: " + compInfor.TotalPhysicalMemory / 1024 / 1024 / 1024 + "GB");
                        sw.WriteLine("Total Virtual Memory: " + compInfor.TotalVirtualMemory / 1024 / 1024 / 1024 + "GB");
                        sw.WriteLine("Total Processes Running: " + Process.GetProcesses().Length);
                        sw.WriteLine(".NET Framework Version: " + Get45PlusFromRegistry());
                        sw.WriteLine("-- End Workstation Information --");
                        sw.WriteLine(" ");
                        sw.WriteLine("\nNote: For more information on Workstation and Network Data, please run the OSCinfo.bat utility as published in Answer 2412 or enable the Add-In ServerProperty WorkstationInfo.");
                }
        }
        catch (Exception e)
        {                
                MessageBox.Show(e.Message);
        }
}

private static string Get45PlusFromRegistry()
{
        const string subkey = @"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\";
        string version;

        using (var ndpKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(subkey))
        {
                version = ndpKey?.GetValue("Release") != null ? CheckFor45PlusVersion((int)ndpKey.GetValue("Release")) : ".NET Framework Version 4.5 or later is not detected.";
        }

        return version;
}

private static string CheckFor45PlusVersion(int releaseKey)
{
        if (releaseKey >= 460798)
                return "4.7 or later";
        if (releaseKey >= 394802)
                return "4.6.2";
        if (releaseKey >= 394254)
                return "4.6.1";
        if (releaseKey >= 393295)
                return "4.6";
        if (releaseKey >= 379893)
                return "4.5.2";
        if (releaseKey >= 378675)
                return "4.5.1";
        return releaseKey >= 378389 ? "4.5" : "No 4.5 or later version detected";
}
 

Otherwise, the OSCInfo.bat file is extracted and will run using the agent session information, without to fill out information in the bat file.

private static void ExtractEmbeddedResource(string nameSpace, string outDirectory, string internalFilePath, string resourceName)
{
        try
        {                
                var assembly = Assembly.GetCallingAssembly();

                using (var s = assembly.GetManifestResourceStream(nameSpace + "." + (internalFilePath == "" ? "" : internalFilePath + ".") + resourceName))
                        if (s != null)
                                using (var r = new BinaryReader(s))
                                using (var fs = new FileStream(outDirectory + "\\" + resourceName, FileMode.OpenOrCreate))
                                using (var w = new BinaryWriter(fs))
                                        w.Write(r.ReadBytes((int)s.Length));
        }
        catch (Exception e)
        {                
                MessageBox.Show(e.Message);
        }
}

private static void RunOsvCInfo()
{
        try
        {
                var proc = new Process
                {
                        StartInfo =
                        {
                                FileName = _rootpath + "\\OSvCinfo.bat",
                                Arguments =
                                        _osvcLoggedin + " " + _osvcInterface + " " + _hostname + " " + _path + "\\" +
                                        DateTime.Now.ToString("MMddyyyyHHmmss") + "_OSvCInfo.txt",
                                WindowStyle = ProcessWindowStyle.Normal,
                                CreateNoWindow = true
                        }
                };
                proc.Start();
                proc.WaitForExit();

                CompressAndNotify();
        }
        catch (Exception e)
        {                
                MessageBox.Show(e.Message);
        }
}
 

Finally, the information is compressed and the result location will be shown to the agent.

private static void CompressAndNotify()
{
        try
        {
                ZipFile.CreateFromDirectory(_path, _path + ".zip", CompressionLevel.Fastest, true);
                Directory.Delete(_path, true);

                var strLocation = _path + ".zip";

                using (var frmCompletionForm = new FrmCompletionForm(strLocation, _finalnotification))
                {
                        frmCompletionForm.ShowDialog();
                }

        }
        catch (Exception e)
        {                
                MessageBox.Show(e.Message);
        }
}
 

Leave a comment and let us know what you think, or if you have questions. We’d also love to hear other ways that you’ve simplified or improved troubleshooting agent issues.
