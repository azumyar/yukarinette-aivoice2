using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace Yarukizero.Net.Yularinette.AiVoice2 {
	internal class Connect : IDisposable {
		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		static extern int RegisterWindowMessage(string lpString);

		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		private static extern nint PostMessage(nint hwnd, int msg, nint wp, nint lp);
		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		private static extern nint SendMessage(nint hwnd, int msg, nint wp, nint lp);

		[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
		private static extern nint LoadLibrary(string lpLibFileName);
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
		static extern nint GetProcAddress(nint hModule, string lpProcName);

		[DllImport("user32.dll")]
		private static extern bool GetClientRect(nint hWnd, out RECT lpRect);
		[DllImport("user32.dll")]
		private static extern bool GetWindowRect(nint hWnd, out RECT lpRect);
		[DllImport("user32.dll")]
		private static extern bool SetWindowPos(nint hWnd, nint hWndInsertAfter, int x, int y, int cx, int cy, int uFlags);

		[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
		private static extern int GetPrivateProfileInt(string lpAppName, string lpKeyName, int nDefault, string lpFileName);

		[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
		private static extern nint CreateFileMapping(nint hFile, nint lpFileMappingAttributes, int flProtect, int dwMaximumSizeHigh, int dwMaximumSizeLow, string lpName);
		[DllImport("kernel32.dll")]
		private static extern nint MapViewOfFile(nint hFileMappingObject, int dwDesiredAccess, int dwFileOffsetHigh, int dwFileOffsetLow, nint dwNumberOfBytesToMap);
		[DllImport("kernel32.dll")]
		private static extern nint UnmapViewOfFile(nint hFileMappingObject);
		[DllImport("kernel32.dll")]
		private static extern nint CloseHandle(nint hObject);
		[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
		private static extern nint lstrcpy(nint str1, string str2);
		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		private static extern nint FindWindowEx(
			nint hWndParent,
			nint hWndChildAfter,
			string lpszClass,
			string lpszWindow);

		private const int WM_KILLFOCUS = 0x0008; 
		private const int WM_LBUTTONDOWN = 0x201;
		private const int WM_LBUTTONUP = 0x202;
		private const int WM_KEYDOWN = 0x0100;
		private const int WM_KEYUP = 0x0101;
		private const int WM_IME_CHAR = 0x286;
		private const int MK_LBUTTON = 0x01;
		private const int VK_HOME = 0x24;
		private const int VK_DELETE = 0x2E;
		private const int VK_SPACE = 0x20;
		private const int VK_F5 = 0x74;

		private const int PAGE_READWRITE = 0x04;
		private const int FILE_MAP_WRITE = 0x00000002;

		[StructLayout(LayoutKind.Sequential)]
		struct POINT {
			public int x;
			public int y;
		}
		[StructLayout(LayoutKind.Sequential)]
		struct RECT {
			public int left;
			public int top;
			public int right;
			public int bottom;
		}
		private readonly string MapNameCapture = "yarukizero-net-yukarinette.audio-capture";

		class MessageWindow : System.Windows.Window {
			private readonly Connect con;
			private System.Windows.Interop.HwndSource source;

			public IntPtr Handle { get; }

			public MessageWindow(Connect con) {
				this.con = con;
				this.ShowInTaskbar = false;
				this.Opacity = 0;
				this.Title = "yukarinette-aivoice2";

				var helper = new System.Windows.Interop.WindowInteropHelper(this);
				this.Handle = helper.EnsureHandle();
				this.source = System.Windows.Interop.HwndSource.FromHwnd(helper.Handle);
				this.source.AddHook(WndProc);
				this.Loaded += async (_, _) => {
					await Task.Delay(100);
					this.Hide();
				};
			}

			private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) {
				if(msg == YACM_STARTUP) {
					this.con.captureWindow = lParam;
				} else if(msg == YACM_CAPTURE) {
					this.con.sync.Set();
					handled = true;
				}
				return IntPtr.Zero;
			}
		}

		private readonly AutoResetEvent sync = new AutoResetEvent(false);

		private int connectionMsg;
		private MessageWindow window;
		private IntPtr formHandle;

		private IntPtr hAiVoice;
		private int voicePeakWidth;

		private readonly string dllPath;
		private readonly string exePath;
		private readonly string iniPath;
		private readonly int defaultWaitSec = 50;

		private Process captureProcess;
		private nint captureWindow;
		private const int WM_APP = 0x8000;
		private const int YACM_STARTUP = WM_APP + 1;
		private const int YACM_SHUTDOWN = WM_APP + 2;
		private const int YACM_GETSTATE = WM_APP + 4;
		private const int YACM_CAPTURE = WM_APP + 5;

		public Connect() {
			this.exePath = Path.Combine(
				AppDomain.CurrentDomain.BaseDirectory,
				"Plugins",
				"Yarukizero.Net.AiVoice2",
				"audio-capture.exe");
			this.iniPath = Path.Combine(
				AppDomain.CurrentDomain.BaseDirectory,
				"Plugins",
				"Yarukizero.Net.AiVoice2",
				"aivoice2.ini");

			this.window = new MessageWindow(this);
			this.window.Show();
			this.formHandle = this.window.Handle;
		}

		public void Dispose() {
			this.window?.Close();
		}

		public bool BeginCapture() {
			return true;
		}

		public bool BeginAiVoice() {
			var p = System.Diagnostics.Process.GetProcesses().Where(x => {
				try {
					return x.MainModule?.ModuleName?.ToLower() == "aivoice.exe";
				}
				catch(Exception e) when((e is Win32Exception) || (e is InvalidOperationException)) {
					return false;
				}
			}).FirstOrDefault();
			if(p == null) {
				return false;
			}
			
			this.hAiVoice = p.MainWindowHandle;
			GetWindowRect(this.hAiVoice, out var rc);
			SetWindowPos(this.hAiVoice, IntPtr.Zero, rc.left, rc.top, 1152, 720, 0);
			GetClientRect(this.hAiVoice, out rc);
			this.voicePeakWidth = rc.right - rc.left;
			// キーボードフォーカス握るウインドウに差し替え
			this.hAiVoice = FindWindowEx(p.MainWindowHandle, 0, "FLUTTERVIEW", "FLUTTERVIEW");

			if(this.captureProcess == null) {
				var hMapObj = CreateFileMapping(
					(IntPtr)(-1), IntPtr.Zero, PAGE_READWRITE,
					0, 8 * 2,
					MapNameCapture);
				var ptr = MapViewOfFile(hMapObj, FILE_MAP_WRITE, 0, 0, IntPtr.Zero);
				Marshal.WriteIntPtr(ptr, 0, (IntPtr)p.Id);
				Marshal.WriteIntPtr(ptr, 8, this.window.Handle);
				UnmapViewOfFile(ptr);
				this.captureProcess = Process.Start(this.exePath);
			}
			return true;
		}

		public void EndCaptureAiVoice() {
			if(this.captureProcess != null) {
				PostMessage(this.captureWindow, YACM_SHUTDOWN, 0, 0);
				this.captureProcess.Dispose();
				this.captureProcess = null;
			}
		}

		public void Speech(string text) {
			static void click(IntPtr hwnd, int x, int y) {
				var pos = x | (y << 16);
				PostMessage(hwnd, WM_LBUTTONDOWN, (IntPtr)MK_LBUTTON, (IntPtr)pos);
				PostMessage(hwnd, WM_LBUTTONUP, IntPtr.Zero, (IntPtr)pos);
			}

			static void keyboard(IntPtr hwnd, int keycode) {
				PostMessage(hwnd, WM_KEYDOWN, (IntPtr)keycode, (IntPtr)(0x000000001));
				PostMessage(hwnd, WM_KEYUP, (IntPtr)keycode, (IntPtr)0xC00000001);
			}


			// キャプチャ待機
			this.sync.Reset();
			PostMessage(
				this.captureWindow,
				YACM_CAPTURE,
				0,
				0);

			// 文字入力
			click(this.hAiVoice, 380, 185);
			Thread.Sleep(100);
			foreach(var c in text) {
				int WM_CHAR = 0x0102;
				SendMessage(this.hAiVoice, WM_CHAR, (IntPtr)c, IntPtr.Zero);
			}
			// 逐次変換されるぽいので固定で1秒待つ
			Thread.Sleep(1000);

			// 再生
			click(this.hAiVoice, 475, 45);
			{
				var waitSec = GetPrivateProfileInt(
					"plugin",
					"waittime_speech",
					this.defaultWaitSec,
					this.iniPath);
				this.sync.WaitOne(waitSec * 1000); // フリーズ防止のためデフォルト50秒で解除する
			}

			// 後片付け
			click(this.hAiVoice, 380, 185);
			Thread.Sleep(100);
			if(!string.IsNullOrEmpty(text)) {
				// 残ることがあるらしいので3週Deleteを打つ
				for(var i = 0; i < 3; i++) {
					keyboard(this.hAiVoice, VK_HOME);
					foreach(var _ in text) {
						keyboard(this.hAiVoice, VK_DELETE);
					}
				}
			}
		}
	}
}