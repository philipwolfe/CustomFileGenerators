using System;
using System.Collections.Generic;
$if$ ($targetframeworkversion$ >= 3.5)using System.Linq;
$endif$using System.Text;

namespace $rootnamespace$
{
	public partial class $safeitemrootname$
	{
		//Note: See http://CustomFileGenerators.codeplex.com for the latest information

		//Uncomment the examples below and save the file to see the observable files generated.
		//The properties are generated in the "g" file as a subitem of this file.
		//There is a custom tool specified in the properties page for this file named "ObsPropGen".

		//[GenerateObservableProperty]
		//private string _name;
		//
		//[GenerateObservableProperty("BirthDate")]
		//private DateTime _dateOfBirth;
	}

	//There are two ways of referencing the "GenerateObservableProperty" attribute
	// 1. Uncomment the code below and compile it into your solution.
	// 2. Reference the attribute from the assembly CustomFileGenerators.ObservablePropertyAttributes.dll

	//[AttributeUsage(AttributeTargets.Field)]
	//public class GenerateObservablePropertyAttribute : Attribute
	//{
	//    public GenerateObservablePropertyAttribute()
	//    {}
	//
	//    public GenerateObservablePropertyAttribute(string propertyName)
	//    {}
	//}
}