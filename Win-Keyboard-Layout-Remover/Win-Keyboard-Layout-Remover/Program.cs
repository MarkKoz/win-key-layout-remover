using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Management.Automation;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Text;
using System.Threading;
using System.IO;
using System.Security.Cryptography;

namespace Win_Keyboard_Layout_Remover
{
	public class Program
	{
		public static void Logger(object str)
		{
			str = str.ToString();
			string d = DateTime.Now.ToString(CultureInfo.InvariantCulture);
			File.AppendAllText("log.log", $"[{d}]\t{str}\n");
			Console.WriteLine($"[{d}] {str}");
		}

		static void Main(string[] args)
		{
			PowerShellExecutor ps = new PowerShellExecutor();

			dynamic list = ps.ExecuteSynchronously("Get-WinUserLanguageList", false, null, null).BaseObject;

			while (true)
			{
				dynamic newList = ps.ExecuteSynchronously("Get-WinUserLanguageList", false, null, null).BaseObject;
				dynamic replaceList = ps.ExecuteSynchronously("Get-WinUserLanguageList", false, null, null).BaseObject;
				int layoutID = GetSet.getCurrentKeyboardLayout();

				if (list.Count != newList.Count)
				{
					Logger($"New language was detected while window \"{GetSet.getActiveWindowTitle()}\" was active.");
					for (int i = 0; i < newList.Count; i++)
					{
						var newItem = newList[i];
						Boolean found = false;

						foreach (var oldItem in list)
						{
							if (newItem.LanguageTag == oldItem.LanguageTag)
							{
								found = true;
							}
						}
						if (!found)
						{
							replaceList.RemoveAt(i);
						}
					}
					ps.ExecuteSynchronously("param($finalList) Set-WinUserLanguageList($finalList) -Force", true, "finalList", replaceList);
					Logger("New language removed.");
					GetSet.setInputMethod(layoutID);
				}
				list = ps.ExecuteSynchronously("Get-WinUserLanguageList", false, null, null).BaseObject;
				Thread.Sleep(1000);
			}
		}
	}

	public class GetSet
	{

		[DllImport("user32.dll")]
		static extern IntPtr GetForegroundWindow();

		[DllImport("user32.dll")]
		static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

		[DllImport("user32.dll")]
		static extern uint GetWindowThreadProcessId(IntPtr hwnd, IntPtr proccess);

		[DllImport("user32.dll")]
		static extern IntPtr GetKeyboardLayout(uint thread);

		[DllImport("user32.dll")]
		public static extern IntPtr PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

		[DllImport("user32.dll")]
		private static extern IntPtr LoadKeyboardLayout(string pwszKLID, uint Flags);

		public static int getCurrentKeyboardLayout()
		{
			try
			{
				IntPtr foregroundWindow = GetForegroundWindow();
				uint foregroundProcess = GetWindowThreadProcessId(foregroundWindow, IntPtr.Zero);
				int keyboardLayout = GetKeyboardLayout(foregroundProcess).ToInt32() & 0xFFFF;
				return keyboardLayout;
			}
			catch (Exception)
			{
				Program.Logger($"Invalid locale ID. Defaulting to {getCultureInfo(1053).DisplayName}.");
				return 1053;
			}
		}

		public static CultureInfo getCultureInfo(int layoutID)
		{
			return new CultureInfo(layoutID);
		}

		public static string getActiveWindowTitle()
		{
			const int nChars = 256;
			StringBuilder Buff = new StringBuilder(nChars);
			IntPtr handle = GetForegroundWindow();

			if (GetWindowText(handle, Buff, nChars) > 0)
			{
				return Buff.ToString();
			}
			return null;
		}

		public static void setInputMethod(int layoutID)
		{
			switch (layoutID)
			{
				case 1053:
					PostMessage(GetForegroundWindow(), 0x0050, (IntPtr)0, LoadKeyboardLayout("A000041D", 1));
					Program.Logger($"Active language set to {getCultureInfo(layoutID).DisplayName}.");
					break;
				case 1049:
					PostMessage(GetForegroundWindow(), 0x0050, (IntPtr)0, LoadKeyboardLayout("A0000419", 1));
					Program.Logger($"Active language set to {getCultureInfo(layoutID).DisplayName}.");
					break;
				default:
					PostMessage(GetForegroundWindow(), 0x0050, (IntPtr)0, LoadKeyboardLayout("A000041D", 1));
					Program.Logger($"Error setting the active language. Defaulting to {getCultureInfo(1053).DisplayName}.");
					break;
			}
		}
	}

	class PowerShellExecutor
	{
		public PSObject ExecuteSynchronously(string script, Boolean b, string p, object pVal)
		{
			using (PowerShell PowerShellInstance = PowerShell.Create())
			{
				PowerShellInstance.AddScript(script);

				if (b)
				{
					PowerShellInstance.AddParameter(p, pVal);
				}
				Collection<PSObject> PSOutput = PowerShellInstance.Invoke();

				foreach (PSObject outputItem in PSOutput)
				{
					if (outputItem != null)
					{
						return outputItem;
					}
				}
			}
			return null;
		}
	}
}
