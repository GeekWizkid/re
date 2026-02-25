using System;
using System.Runtime.InteropServices;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using System.Resources;
using System.Reflection;
using System.IO;
using System.Globalization;
using System.CodeDom.Compiler;
using Microsoft.CSharp;
using System.Windows.Forms;

namespace Decompiler
{
	/// <summary>The object for implementing an Add-in.</summary>
	/// <seealso class='IDTExtensibility2' />
	public class Connect : IDTExtensibility2, IDTCommandTarget
	{
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr LoadLibrary(string libname);

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        private static extern IntPtr GetProcAddress(IntPtr hModule, string procName);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FreeLibrary(IntPtr hModule);

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate void ConnectDelegate(int hWnd, IntPtr pDTE, IntPtr pDebugger);

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate void DisconnectDelegate();

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int WriteAddrDelegate(string message);

        // NEW: window notifier to trigger "go to source" after jump
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate void SetNotifyWindowDelegate(IntPtr hwnd);

        private DTE2 _applicationObject;
        private AddIn _addInInstance;
        private EnvDTE.Debugger _debugger;
        private DebuggerEvents _debuggerEvents;

        private IntPtr hModule;
        private ConnectDelegate _Connect;
        private DisconnectDelegate _Disconnect;
        private WriteAddrDelegate _WriteAddr;
        private SetNotifyWindowDelegate _SetNotifyWindow;

        private JumpNotifyWindow _jumpWindow;

        private string currentDirectory;
        private string LogFilePath;

        // Message from Pipe.dll after it "typed" the address into Disassembly
        private sealed class JumpNotifyWindow : NativeWindow
        {
            internal const int WM_IDA_JUMP_DONE = 0x8000 + 0x123; // WM_APP + 0x123
            private readonly Connect _owner;
            private readonly Timer _timer;
            private int _tries;

            public JumpNotifyWindow(Connect owner)
            {
                _owner = owner;

                CreateHandle(new CreateParams());

                _timer = new Timer();
                _timer.Interval = 60; // ms
                _timer.Tick += new EventHandler(OnTimerTick);
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_IDA_JUMP_DONE)
                {
                    _tries = 0;
                    _timer.Stop();
                    _timer.Start();
                }

                base.WndProc(ref m);
            }

            private void OnTimerTick(object sender, EventArgs e)
            {
                _tries++;

                bool done = _owner.TryGoToSourceOnce();
                if (done || _tries >= 10)
                {
                    _timer.Stop();
                }
            }

            public void Dispose()
            {
                try
                {
                    _timer.Stop();
                    _timer.Tick -= new EventHandler(OnTimerTick);
                    _timer.Dispose();
                }
                catch
                {
                }

                try
                {
                    DestroyHandle();
                }
                catch
                {
                }
            }
        }

        public Connect()
		{
		}

        private void Log(string message)
        {
            try
            {
                if (LogFilePath == null || LogFilePath.Length == 0)
                    return;

                string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture) +
                              " " + message + Environment.NewLine;
                File.AppendAllText(LogFilePath, line);
            }
            catch
            {
            }
        }

        private bool TryExecuteCommand(string command)
        {
            try
            {
                _applicationObject.ExecuteCommand(command, "");
                Log("ExecuteCommand OK: " + command);
                return true;
            }
            catch (Exception ex)
            {
                Log("ExecuteCommand FAIL: " + command + " | " + ex.Message);
                return false;
            }
        }

                // Called after Pipe.dll reports that it performed jump-to-address in Disassembly
        internal bool TryGoToSourceOnce()
        {
            try
            {
                if (_applicationObject == null)
                    return true;

                // Usually these commands are only available in Break mode
                try
                {
                    if (_applicationObject.Debugger != null &&
                        _applicationObject.Debugger.CurrentMode != dbgDebugMode.dbgBreakMode)
                    {
                        return false;
                    }
                }
                catch
                {
                    // ignore
                }

                // First try (may not exist on VS2005)
                if (TryExecuteCommand("Debug.GoToSourceCode"))
                    return true;

                // Fallback that exists in VS2005: toggles between disassembly/source.
                // Since we just jumped to Disassembly, one toggle should go to source (if mapping exists).
                if (TryExecuteCommand("Debug.ToggleDisassembly"))
                    return true;
            }
            catch (Exception ex)
            {
                Log("TryGoToSourceOnce error: " + ex.Message);
            }
            return false;
        }

private static long ParseHexInt64(string s)
        {
            if (s == null)
                return 0;

            s = s.Trim();
            if (s.StartsWith("0x") || s.StartsWith("0X"))
                s = s.Substring(2);

            // Some VS values may include backticks or separators; keep only hex
            // (very conservative)
            int i = 0;
            while (i < s.Length)
            {
                char c = s[i];
                bool isHex = (c >= '0' && c <= '9') ||
                             (c >= 'a' && c <= 'f') ||
                             (c >= 'A' && c <= 'F');
                if (!isHex)
                    break;
                i++;
            }
            if (i != s.Length)
                s = s.Substring(0, i);

            return Convert.ToInt64(s, 16);
        }

		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			_applicationObject = (DTE2)application;
			_addInInstance = (AddIn)addInInst;

            //if (connectMode == ext_ConnectMode.ext_cm_UISetup)
            if (connectMode == ext_ConnectMode.ext_cm_Startup || connectMode == ext_ConnectMode.ext_cm_AfterStartup)
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                string assemblyPath = assembly.Location;
                currentDirectory = Path.GetDirectoryName(assemblyPath) + @"\";

                LogFilePath = currentDirectory + "Decompiler.log";
                Log("Add-in OnConnection (mode=" + connectMode.ToString() + ")");

                object[] contextGUIDS = new object[] { };
                Commands2 commands = (Commands2)_applicationObject.Commands;
                string toolsMenuName;

                try
                {
                    //If you would like to move the command to a different menu, change the word "Tools" to the 
                    //  English version of the menu. This code will take the culture, append on the name of the menu
                    //  then add the command to that menu. You can find a list of all the top-level menus in the file
                    //  CommandBar.resx.
                    ResourceManager resourceManager = new ResourceManager("Decompiler.CommandBar", Assembly.GetExecutingAssembly());
                    CultureInfo cultureInfo = new System.Globalization.CultureInfo(_applicationObject.LocaleID);
                    string resourceName = String.Concat(cultureInfo.TwoLetterISOLanguageName, "Tools");
                    toolsMenuName = resourceManager.GetString(resourceName);
                }
                catch
                {
                    //We tried to find a localized version of the word Tools, but one was not found.
                    //  Default to the en-US word, which may work for the current culture.
                    toolsMenuName = "Tools";
                }

                //Place the command on the tools menu.
                //Find the MenuBar command bar, which is the top-level command bar holding all the main menu items:
                Microsoft.VisualStudio.CommandBars.CommandBar menuBarCommandBar = ((Microsoft.VisualStudio.CommandBars.CommandBars)_applicationObject.CommandBars)["MenuBar"];

                //Find the Tools command bar on the MenuBar command bar:
                CommandBarControl toolsControl = menuBarCommandBar.Controls[toolsMenuName];
                CommandBarPopup toolsPopup = (CommandBarPopup)toolsControl;

                //This try/catch block can be duplicated if you wish to add multiple commands to be handled by your Add-in,
                //  just make sure you also update the QueryStatus/Exec method to include the new command names.
                try
                {
                    //Add a command to the Commands collection:
                    Command command = commands.AddNamedCommand2(_addInInstance, "Decompiler", "Decompiler", "Executes the command for Decompiler", true, 59, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);

                    //Add a control for the command to the tools menu:
                    if ((command != null) && (toolsPopup != null))
                    {
                        command.AddControl(toolsPopup.CommandBar, 1);
                    }
                }
                catch (System.ArgumentException)
                {
                    //If we are here, then the exception is probably because a command with that name
                    //  already exists. If so there is no need to recreate the command and we can 
                    //  safely ignore the exception.
                }

                _debugger = _applicationObject.Debugger;
                _debuggerEvents = _applicationObject.Events.DebuggerEvents;
                _debuggerEvents.OnContextChanged += DebugEvents_OnContextChanged;

                Window win = _applicationObject.MainWindow;

                IntPtr pDTE = Marshal.GetIDispatchForObject(_applicationObject);
                IntPtr pDebugger = Marshal.GetIDispatchForObject(_applicationObject.Debugger);

                hModule = LoadLibrary(currentDirectory + "Pipe.dll");
                if (hModule == IntPtr.Zero)
                {
                    Log("LoadLibrary(Pipe.dll) failed. Win32Error=" + Marshal.GetLastWin32Error().ToString());
                    return;
                }

                // _Connect
                IntPtr functionPointer = GetProcAddress(hModule, "_Connect");
                if (functionPointer == IntPtr.Zero)
                {
                    Log("GetProcAddress(_Connect) failed. Win32Error=" + Marshal.GetLastWin32Error().ToString());
                    return;
                }
                _Connect = (ConnectDelegate)Marshal.GetDelegateForFunctionPointer(functionPointer, typeof(ConnectDelegate));

                // _Disconnect
                functionPointer = GetProcAddress(hModule, "_Disconnect");
                if (functionPointer == IntPtr.Zero)
                {
                    Log("GetProcAddress(_Disconnect) failed. Win32Error=" + Marshal.GetLastWin32Error().ToString());
                    return;
                }
                _Disconnect = (DisconnectDelegate)Marshal.GetDelegateForFunctionPointer(functionPointer, typeof(DisconnectDelegate));

                // _WriteAddr
                functionPointer = GetProcAddress(hModule, "_WriteAddr");
                if (functionPointer == IntPtr.Zero)
                {
                    Log("GetProcAddress(_WriteAddr) failed. Win32Error=" + Marshal.GetLastWin32Error().ToString());
                    return;
                }
                _WriteAddr = (WriteAddrDelegate)Marshal.GetDelegateForFunctionPointer(functionPointer, typeof(WriteAddrDelegate));

                // NEW: optional notifier
                functionPointer = GetProcAddress(hModule, "_SetNotifyWindow");
                if (functionPointer != IntPtr.Zero)
                {
                    _SetNotifyWindow = (SetNotifyWindowDelegate)Marshal.GetDelegateForFunctionPointer(functionPointer, typeof(SetNotifyWindowDelegate));
                    Log("_SetNotifyWindow found");
                }
                else
                {
                    Log("GetProcAddress(_SetNotifyWindow) missing (old Pipe.dll?)");
                }

                // Start pipe server + connect
                _Connect(win.HWnd, pDTE, pDebugger);

                // Register notify window
                try
                {
                    if (_SetNotifyWindow != null)
                    {
                        _jumpWindow = new JumpNotifyWindow(this);
                        _SetNotifyWindow(_jumpWindow.Handle);
                        Log("Notify HWND registered");
                    }
                }
                catch (Exception ex)
                {
                    Log("Register notify HWND failed: " + ex.Message);
                }
            }
		}

		/// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
		/// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		{
            try
            {
                // Unsubscribe events
                if (_debuggerEvents != null)
                {
                    _debuggerEvents.OnContextChanged -= DebugEvents_OnContextChanged;
                    _debuggerEvents = null;
                }

                // Unregister notify window
                if (_SetNotifyWindow != null)
                {
                    try
                    {
                        _SetNotifyWindow(IntPtr.Zero);
                    }
                    catch { }
                    _SetNotifyWindow = null;
                }

                if (_jumpWindow != null)
                {
                    _jumpWindow.Dispose();
                    _jumpWindow = null;
                }

                if (_Disconnect != null)
                {
                    _Disconnect();
                    _Disconnect = null;
                }
Log("Add-in OnDisconnection (" + disconnectMode.ToString() + ")");
            }
            catch
            {
            }
		}

		/// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />		
		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref Array custom)
		{
		}

		/// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref Array custom)
		{
		}

		/// <summary>Implements the QueryStatus method of the IDTCommandTarget interface. This is called when the command's availability is updated</summary>
		/// <param term='commandName'>The name of the command to determine state for.</param>
		/// <param term='neededText'>Text that is needed for the command.</param>
		/// <param term='status'>The state of the command in the user interface.</param>
		/// <param term='commandText'>Text requested by the neededText parameter.</param>
		/// <seealso class='Exec' />
		public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
		{
			if(neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
			{
				if(commandName == "Decompiler.Connect.Decompiler")
				{
					status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported|vsCommandStatus.vsCommandStatusEnabled;
					return;
				}
			}
		}

		/// <summary>Implements the Exec method of the IDTCommandTarget interface. This is called when the command is invoked.</summary>
		/// <param term='commandName'>The name of the command to execute.</param>
		/// <param term='executeOption'>Describes how the command should be run.</param>
		/// <param term='varIn'>Parameters passed from the caller to the command handler.</param>
		/// <param term='varOut'>Parameters passed from the command handler to the caller.</param>
		/// <param term='handled'>Informs the caller if the command was handled or not.</param>
		/// <seealso class='Exec' />
        private Assembly MyResolveEventHandler(object sender, ResolveEventArgs args)
        {
            // Проверка имени запрашиваемой сборки
            if (args.Name.Contains("Interop.HYSYS"))
            {
                string path = currentDirectory + "Interop.HYSYS.dll";
                return Assembly.LoadFrom(path);
            }
            return null;
        }

        private void ExecuteDecompiler()
        {
            try
            {
                string code = File.ReadAllText(currentDirectory + "Script.cs");
                CSharpCodeProvider provider = new CSharpCodeProvider();
                CompilerParameters parameters = new CompilerParameters();

                parameters.ReferencedAssemblies.Add("System.dll");
                parameters.ReferencedAssemblies.Add(currentDirectory + "Interop.HYSYS.dll");
                parameters.GenerateInMemory = true;

                CompilerResults results = provider.CompileAssemblyFromSource(parameters, code);
                if (!results.Errors.HasErrors)
                {
                    AppDomain currentDomain = AppDomain.CurrentDomain;
                    currentDomain.AssemblyResolve += new ResolveEventHandler(MyResolveEventHandler);
                    Assembly assembly = results.CompiledAssembly;
                    Type type = assembly.GetType("Script");
                    object obj = Activator.CreateInstance(type);
                    MethodInfo methodInfo = type.GetMethod("Run");
                    methodInfo.Invoke(obj, null);
                }
            }
            catch (Exception)
            {
            }
        }

        public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
		{
			handled = false;
			if(executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
			{
				if(commandName == "Decompiler.Connect.Decompiler")
				{
                    System.Threading.Thread newThread = new System.Threading.Thread(new System.Threading.ThreadStart(ExecuteDecompiler));
                    newThread.Start();
                    handled = true;
					return;
				}
			}
		}

        private void DebugEvents_OnContextChanged(EnvDTE.Process NewProcess, Program NewProgram, EnvDTE.Thread NewThread, EnvDTE.StackFrame NewStackFrame)
        {
            if (NewStackFrame != null)
            {
                EnvDTE.Expression expression = null;

                // x64 first
                try
                {
                    expression = _debugger.GetExpression("rip", true, 1000);
                }
                catch
                {
                }

                // x86 fallback
                if (expression == null || !expression.IsValidValue)
                {
                    try
                    {
                        expression = _debugger.GetExpression("eip", true, 1000);
                    }
                    catch
                    {
                    }
                }

                if (expression != null && expression.IsValidValue)
                {
                    string addr = expression.Value;

                    try
                    {
                        if (NewStackFrame != _applicationObject.Debugger.CurrentThread.StackFrames.Item(1))
                        {
                            long lAddr = ParseHexInt64(addr);
                            lAddr -= 5;
                            addr = "0x" + lAddr.ToString("X");
                        }
                    }
                    catch
                    {
                    }

                    addr += '\n';
                    try
                    {
                        if (_WriteAddr != null)
                            _WriteAddr(addr);
                    }
                    catch
                    {
                    }
                }
            }
        }
	}
}
