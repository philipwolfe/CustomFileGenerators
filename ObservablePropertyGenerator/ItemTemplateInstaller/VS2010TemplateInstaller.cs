using System;
using System.Collections;
using System.ComponentModel;
using System.IO;
using System.Security.Permissions;
using Microsoft.Win32;

namespace ItemTemplateInstaller
{
	[RunInstaller(true)]
	public class VS2010TemplateInstaller : VSTemplateInstallerBase
	{
		private const string LOG_FILE_NAME = "VS2010TemplateInstaller.log";

		//public VS2010TemplateInstaller() : base(LOG_FILE_NAME)
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
		
		protected override string InstallDir
		{
			get
			{
				if (Environment.Is64BitOperatingSystem)
				{
					return Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\10.0", "InstallDir", string.Empty).ToString();
				}
				else
				{
					return Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\10.0", "InstallDir", string.Empty).ToString();
				}
			}
		}

		protected override string ProductTitle
		{
			get { return "Visual Studio 2010"; }
		}

		protected override string CustomActionHeader
		{
			get
			{
				return GetType().Name;
			}
		}

		protected override int VSVersionMajor
		{
			get { return 10; }
		}

		protected override string DevEnvLocation
		{
			get { return Path.Combine(InstallDir, "devenv.exe"); }
		}
	}
}
