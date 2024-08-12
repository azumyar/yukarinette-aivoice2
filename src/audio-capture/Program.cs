using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Yarukizero.Net.Yularinette.VoicePeakConnect;

namespace Yarukizero.Net.Yularinette.AudioCapture;
internal static class Program {
	class MessageForm : Form {
		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		public static extern nint PostMessage(nint hwnd, int msg, nint wParam, nint lParam);

		private const int WM_APP = 0x8000;
		private const int YACM_STARTUP = WM_APP + 1;
		private const int YACM_SHUTDOWN = WM_APP + 2;
		private const int YACM_GETSTATE = WM_APP + 4;
		private const int YACM_CAPTURE = WM_APP + 5;

		private const int YACSTATE_NONE = 0;
		private const int YACSTATE_FAIL = 1;
		private const int YACSTATE_INITILIZED = 3;

		private nint reciveWnd;
		private int targetProcess;
		private bool isInit = false;
		private ApplicationCapture? capture;

		public MessageForm() {
			this.FormBorderStyle = FormBorderStyle.None;
			this.Opacity = 0;
			this.ControlBox = false;
		}

		protected override void WndProc(ref Message m) {
			switch(m.Msg) {
			case YACM_STARTUP:
				this.targetProcess = m.WParam.ToInt32();
				this.reciveWnd = m.LParam;
				Task.Run(async () => {
					this.capture = await ApplicationCapture.Get(this.targetProcess);
					this.isInit = true;
				});
				break;
			case YACM_SHUTDOWN:
				this.Close();
				break;
			case YACM_GETSTATE:
				if(isInit) {
					m.Result = this.capture switch {
						null => YACSTATE_FAIL,
						_ => YACSTATE_INITILIZED,
					};
				} else {
					m.Result = YACSTATE_NONE;
				}
				break;
			case YACM_CAPTURE:
				if(this.capture != null) {
					var index = m.WParam.ToInt32();
					this.capture.Start();
					Task.Run(() => {
						this.capture.Wait();
						this.capture.Stop();
						PostMessage(reciveWnd, YACM_CAPTURE, index, 0);
					});
				}
				break;
			}

			base.WndProc(ref m);
		}

		protected override async void OnLoad(EventArgs e) {
			base.OnLoad(e);
			ApplicationCapture.UiInitilize();
			await Task.Delay(500);
			this.Hide();
		}

		protected override void OnFormClosed(FormClosedEventArgs e) {
			base.OnFormClosed(e);
			Application.Exit();
		}
	}

	[STAThread]
	static void Main() {
		ApplicationConfiguration.Initialize();
		Application.Run(new MessageForm());
	}
}
