using System;

namespace CustomFileGenerators.ObservablePropertyAttributes
{
	/// <summary>
	/// Place this attribute on fields of a ViewModel class to automatically generate properties
	/// that call the RaisePropertyChanged method.
	/// </summary>
	[AttributeUsage(AttributeTargets.Field)]
	public class GenerateObservablePropertyAttribute : Attribute
	{
		/// <summary>
		/// Generates an Observable Property using the name of the field.
		/// </summary>
		public GenerateObservablePropertyAttribute()
		{}

		/// <summary>
		/// Generates an Observable Property using the specified name.
		/// </summary>
		/// <param name="propertyName"></param>
		public GenerateObservablePropertyAttribute(string propertyName)
		{}
	}
}