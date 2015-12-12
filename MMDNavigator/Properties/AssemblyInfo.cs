using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("SPMMDNavigator")]
[assembly: AssemblyDescription(@"Use this tool to Navigate, Administer, and Export SharePoint 2010 Managed Metadata.

Features includes:
-Navigate the MMD structure, seeing more information than the standard MMD interface (like GUIDs, internal settings, counts, etc.)
-Updating and creating terms, termsets, groups
-Exporting using either a custom format which includes term GUIDs and labels/synonyms
-Exporting using the standard Microsoft format that can be imported using SharePoint MMD
-Importing using a custom XML format, supporting termset GUIDs, term GUIDs, term Labels/Synonyms, Sub Terms, creating or updating Termsets/Terms


Export Options:
1. Microsoft CSV Format:
Export a MMD Termset (one at a time) to the built-in Microsoft SharePoint MMD Import File format.  This does not include Term IDs (GUID) or Term labels/synonyms.

2. Custom CSV Format:
I built this format to make it easy to export MMD data to be easily imported into a database (like SQL Server) or Excel, Access, etc.  It includes Term IDs (GUID) and Term labels/synonyms.  The entire TermStore can be exported, or a Group, or a TermSet.  Additionally, an option to split the synonyms to separate rows when exporting is provided (or the synonyms are added as additional csv columns).

3. Custom Xml Format:
I built this format to make it easy to export the MMD data so that it can be imported by this program.  It exports 1 or more TermSets at a time, and includes all important TermSet, Term, and Label information.  This option is further explained below.


Xml File Import Instructions:

An Xml file can be provided to import that contains 1 or more TermSets and Terms.  Terms can also include synonyms/labels, GUIDs, and other settings described below.  TermSet GUIDs can be included so that existing TermSets are updated instead of created.  Also, if TermSet name is included and found then it will be updated instead of created.  Term GUIDs can be included to reuse existing Terms if the reuse attribute equals true.  If reuse attribute is true, then the reusebranch attribute can be set to true so that the entire Term branch is reused.  If reusebranch is true and the Xml file specifies sub-terms within the current term being imported, then the sub-terms will be ignored since the branch is used.  If reuse attribute is false and the GUID is provided and a matching Term is found, then the term is not reused but will be updated (including the name and other provided attributes).  If the term GUID is not found and/or the Term name is provided and already exists, then it will be updated.  One or more Labels for terms can also be provided, but the Label should be different than the term name.  Terms can contain other terms (add the Xml element after any labels).  This program does not limit the nesting of terms, but SharePoint best practices suggests 7 levels is the maximum supported limit of nesting.

Xml Definition:

termset (contains term elements)
	name=termset name, required (updated when new termset is created, or termset is found with matching GUID)
	id=GUID/empty, optional (leave blank and new termset is created, provide an existing GUID to update the termset)
	description=text, optional
	isavailfortagging=true/false, optional (default is false)
	isopenfortermcreation=true/false, optional (default is false)

	term (can contain term elements and label elements)
		name=term name, required (updated when new term is created, or term is found with matching GUID)
		id=GUID/empty, optional (leave blank and new term is created, when reuse=true and GUID is provided and found then term is reused, if reuse=false and GUID is provided and found then term is updated)
		description=text, optional
		isavailfortagging=true/false, optional (default is false)
		reuse=true/false (when true, requires id to contain valid GUID of existing term, otherwise term is created)
		reusebranch=true/false (when true, requires reuse=true, and reuse conditions described above)

		label (optional, when provided should be direct child of term element, one label element per label/synonym)
			name=text, required (should NOT match term name)


Sample Xml Specification:

<termsets> (with 2 termsets)
	<termset>
		<term> (with 2 labels)
			<label></label>
			<label></label>
		</term>
		<term> (with 1 label, 2 sub-terms, 1 direct descendent term)
			<label></label>
			<term> (with 0 labels, 1 sub-term)
				<term></term>
			</term>
		</term>
	<termset>
	<termset> (with 2 terms)
		<term></term>
		<term></term>
		<term> (with 2 terms)
			<term></term>
			<term></term>
		</term>
	</termset>
</termsets>

")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("B&&R Business Solutions, Authored by Ben Steinhauser")]
[assembly: AssemblyProduct("SPMMDNavigator")]
[assembly: AssemblyCopyright("Copyright  2012")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("2c4d54b5-1558-49ad-9ce9-e3ddfc7bc05e")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers 
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]
