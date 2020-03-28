using System.Reflection;
using System.Runtime.InteropServices;

#if SILVERLIGHT
[assembly: AssemblyTitle("ObservablePropertyAttributes.SL4")]
#else
[assembly: AssemblyTitle("ObservablePropertyAttributes")]
#endif
[assembly: AssemblyDescription("Contains attributes that are used by the ObservablePropertyGenerator tool.")]
#if DEBUG
[assembly: AssemblyConfiguration("DEBUG")]
#else
[assembly: AssemblyConfiguration("RELEASE")]
#endif
[assembly: AssemblyCompany("www.PhilipWolfe.com")]
[assembly: AssemblyProduct("CustomFileGenerators")]
[assembly: AssemblyCopyright("Copyright © Philip Wolfe 2012")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

[assembly: ComVisible(false)]

[assembly: AssemblyVersion("1.3.0.0")]
[assembly: AssemblyFileVersion("1.3.0.0")]