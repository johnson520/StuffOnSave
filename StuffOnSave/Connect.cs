using System;
using System.IO;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using PrefixCSS;

namespace StuffOnSave
{
	/// <summary>The object for implementing an Add-in.</summary>
	/// <seealso class='IDTExtensibility2' />
	public class Connect : IDTExtensibility2
	{
		/// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
// ReSharper disable EmptyConstructor
		public Connect()
		{
		}
// ReSharper restore EmptyConstructor

		/// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			_applicationObject = (DTE2)application;
			_addInInstance = (AddIn)addInInst;
			
            _documentEvents = _applicationObject.Events.DocumentEvents;
            _documentEvents.DocumentSaved += DocumentSaved;


			_CSSPrefixer = new CSSPrefixes(_applicationObject);
		}

	    private void DocumentSaved(Document document)
	    {
// ReSharper disable PossibleNullReferenceException
	        var fileName = Path.GetFileName(document.FullName);
	        var extension = Path.GetExtension(document.FullName);

	        if (extension == ".css" && !fileName.EndsWith(".prefixed.css"))
	        {
				_CSSPrefixer.Add(document.FullName);

				//var exePath = @"C:\Users\Ted\Documents\GitHub\PrefixCSS\PrefixCSS\bin\Debug\PrefixCSS.exe";
				//if (!File.Exists(exePath))
				//	exePath = @"D:\GitHub\PrefixCSS\PrefixCSS\bin\Debug\PrefixCSS.exe";

				//if (File.Exists(exePath))
				//	System.Diagnostics.Process.Start(exePath, document.FullName);
	        }
// ReSharper restore PossibleNullReferenceException

	}

	    /// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		{
		}

		/// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
		public void OnStartupComplete(ref Array custom)
		{
		}

		/// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
		public void OnBeginShutdown(ref Array custom)
		{
		}
		
		private DTE2 _applicationObject;
// ReSharper disable NotAccessedField.Local
		private AddIn _addInInstance;
// ReSharper restore NotAccessedField.Local
	    private DocumentEvents _documentEvents;
		private CSSPrefixes _CSSPrefixer;
	}
}