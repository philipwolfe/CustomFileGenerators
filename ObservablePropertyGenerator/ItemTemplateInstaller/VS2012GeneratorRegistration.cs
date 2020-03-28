using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Security.Permissions;
using Microsoft.Win32;

namespace ItemTemplateInstaller
{
	[RunInstaller(true)]
	public class VS2012GeneratorRegistration : Installer
	{
		private const string ARGUMENT_VS12Setup = "VS12Setup";
		private const string SETUP_VS2012 = "Setup VS2012";
		private const string LOG_FILE_NAME = "VS2012GeneratorRegistration.log";

		//public VS2012GeneratorRegistration() : base(LOG_FILE_NAME)
		//{	
		//}

		[SecurityPermission(SecurityAction.Demand)]
		public override void Install(IDictionary stateSaver)
		{
			InstallInternal(stateSaver);
		}

		[SecurityPermission(SecurityAction.Demand)]
		public override void Rollback(IDictionary stateSaver)
		{
			UninstallInternal(stateSaver);
		}

		[SecurityPermission(SecurityAction.Demand)]
		public override void Uninstall(IDictionary stateSaver)
		{
			UninstallInternal(stateSaver);
		}

		void InstallInternal(IDictionary stateSaver)
		{
			if (!string.IsNullOrEmpty(InstallDir))
			{
				SetupVS();

				stateSaver.Add(SETUP_VS2012, true);
			}
		}

		void UninstallInternal(IDictionary stateSaver)
		{
			if (stateSaver == null)
				return;

			if(stateSaver.Contains(SETUP_VS2012))
			{
				SetupVS();
			}
		}

		private void SetupVS()
		{
			Context.LogMessage(CustomActionHeader + "Running setup on " + ProductTitle);

			var process = Process.Start(DevEnvLocation, "/setup");

			if (process != null)
				process.WaitForExit();

			Context.LogMessage(CustomActionHeader + "Setup complete.");
		}

		string ProductTitle
		{
			get { return "Visual Studio 2012"; }
		}

		string CustomActionHeader
		{
			get
			{
				return GetType().Name;
			}
		}

		string DevEnvLocation
		{
			get { return Path.Combine(InstallDir, "devenv.exe"); }
		}

		string InstallDir
		{
			get
			{
				if (Environment.Is64BitOperatingSystem)
				{
					return Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\11.0", "InstallDir", string.Empty).ToString();
				}
				else
				{
					return Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\11.0", "InstallDir", string.Empty).ToString();
				}
			}
		}

	}
}