<div align="center">

## RECORDSET COPYING USING ADO 2\.5 STREAM OBJECT AND XML \(MDAC 2\.5 REQUIRED\)


</div>

### Description

PROBLEM:

Creating a copy of a Recordset could mean using the Recordset.Clone method sometimes. However, this is not always appropriate because ADO’s Recordset.Clone method creates a Recordset which points to the same data as the original Recordset. This means that you really do not get a physically distinct copy of the original Recorsdet. Consequently, changes to the clone copy could modify the data in the original Recordset.

An alternative solution would be to use a loop to append each record to a new recordset. This, however could end up being an unwieldy solution with potential performance penalties.

SOLUTION:

ADO 2.5 comes with a Stream object that can be used to create physically distinct copies of Recordsets. The article below demonstrates the use of the ADO 2.5 Stream object in copying Recordsets and the difference between the Recordsets created by the Recordset.Clone method and ADO 2.5 Stream object.

Please see attached file.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-07-09 21:50:22
**By**             |[visual\-basic\-data\-mining\.net](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/visual-basic-data-mining-net.md)
**Level**          |Intermediate
**User Rating**    |4.5 (36 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[RECORDSET\_104146792002\_5\_STREAM\_OBJECT\_AND\.zip](https://github.com/Planet-Source-Code/visual-basic-data-mining-net-recordset-copying-using-ado-2-5-stream-object-and-xml-mdac-2-__1-36754/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>PROBLEM</title>
</head>
<body>
<p><font face="Garamond"><b>RECORDSET COPYING USING ADO 2.5 STREAM OBJECT</b><br>
<br>
<b>PROBLEM:</b><br>
Creating a copy of a Recordset could mean using the Recordset.Clone method sometimes. However, this is not always appropriate because ADO’s Recordset.Clone method creates a Recordset which points to the same data as the original Recordset. This means that you really do not get a physically distinct copy of the original Recorsdet. Consequently, changes to the clone copy could modify the data in the original Recordset.<br>
<br>
An alternative solution would be to use a loop to append each record to a new recordset. This, however could end up being an unwieldy solution with potential performance penalties.<br>
<br>
<b>SOLUTION:</b><br>
ADO 2.5 comes with a Stream object that can be used to create physically distinct copies of Recordsets. The article below demonstrates the use of the ADO 2.5 Stream object in copying Recordsets and the difference between the Recordsets created by the Recordset.Clone method and ADO 2.5 Stream object.<br>
<br>
The Clone() Function below demonstrates a typical use of the Recordset.Clone method. <br>
<br>
<font color="#0000FF">Function Clone (ByVal rstSource As Adodb.recordset) As Adodb.Recordset<br>
 'Create a copy of a Recordset using ADO's Recordset.Clone method<br>
<br>
 Dim rstCopy As ADODB.Recordset<br>
<br>
 Set rstCopy = rstSource.Clone<br>
<br>
 Set Clone = rstCopy<br>
<br>
End Function<br>
</font><br>
Since both the initial Recordset and the cloned copy point to the same data structure, data changes like adding and deleting records made to the cloned copy will also take place on the original Recordset and vice versa. So your clone is not really a separate physical copy of the original Recordset.<br>
<br>
<b>How about the ADO 2.5 Stream object?</b><br>
With ADO 2.5, you can create a separate physical Recordset object using the Stream object and XML. It is really very simple as the following Function shows.<br>
<br>
<font color="#0000FF">Function Copy (ByVal rstSource As Adodb.recordset) As Adodb.Recordset<br>
 'Create a copy of a Recordset using AD0 2.5 Stream Object and XML<br>
<br>
 Dim rstCopy As ADODB.Recordset<br>
 Dim objStream As ADODB.Stream<br>
<br>
  'Create a New ADO 2.5 Stream object<br>
  Set objStream = New ADODB.Stream<br>
<br>
  'Save the Recordset to the Stream object in XML format<br>
  rstSource.Save objStream, adPersistXML<br>
<br>
  'Create an exact copy of the saved Recordset from the Stream Object<br>
  Set rstCopy = New ADODB.Recordset<br>
  rstCopy.Open objStream<br>
<br>
  'Close and de-reference the Stream object<br>
  objStream.Close<br>
  Set objStream = Nothing<br>
<br>
  Set Copy = rstCopy <br>
<br>
End Function</font><br>
<br>
The article continued below demonstrates the use of the Recordset.Clone method and the ADO Stream object to copy Recordsets and test the copied Recordsets for references to their original Recordsets.<br>
<br>
Let’s start by creating two Fabricated Recordsets with each containing 3 rows of names. We will also create two copies of the Recordset using the Recordset.Clone method and the ADO Stream object. One name will be deleted from each copied Recordset and finally the original Recordsets will be compared to their copies to see if they were indeed true physically distinct Recordset copies.<br>
<br>
<font color="#0000FF">Sub CreateDistinctRecordsetCopy() <br>
 'Compare ADO's Recordset.Clone and ADO 2.5 Stream object’s Recordset copying methods <br>
<br>
<br>
 Dim rstOne As ADODB.Recordset<br>
 Dim rstTwo As ADODB.Recordset<br>
 Dim rstClone As ADODB.Recordset<br>
 Dim rstCopy As ADODB.Recordset<br>
<br>
<br>
 'Create two Fabricated Recordsets<br>
 Set rstOne = CreateFabricatedRecordset<br>
<br>
 Set rstTwo = CreateFabricatedRecordset<br>
<br>
<br>
 'Create a cloned copy of the Recordset using ADO's Recordset.Clone method<br>
 Set rstClone = Clone(rstOne)<br>
<br>
 'Create a copy of the Recordset using the ADO 2.5 Stream object<br>
 Set rstCopy = Copy(rstTwo)<br>
<br>
<br>
 'Delete a record from both Recordset copies<br>
 rstClone.Delete<br>
 rstCopy.Delete<br>
<br>
<br>
 'If the cloned Recordset and it's original contain the same number of records then ADO Recordset.Clone copies<br>
 'and their original recordsets point to the same data structures and are not completely distinct copies.<br>
<br>
 If (rstOne.RecordCount = rstClone.RecordCount) Then<br>
<br>
  MsgBox "Recordset.Clone Copies Are Not Completely Separate Objects From Their Original Recordsets"<br>
<br>
 Else<br>
<br>
  MsgBox "Recordset.Clone Copies Are Completely Separate Objects From Their Original Recordsets"<br>
<br>
 End If<br>
<br>
<br>
 'If the Recordset copied using ADO 2.5 Stream object and it's original contain differing number of records then<br>
 'ADO 2.5 Stream object Recordset copies are completely distinct copies of the original recordsets.<br>
<br>
 If (rstTwo.RecordCount = rstCopy.RecordCount) Then<br>
<br>
  MsgBox "ADO 2.5 Stream Recordset Copies Are Not Completely Separate Objects From Their Original Recordsets"<br>
<br>
 Else<br>
<br>
  MsgBox "ADO 2.5 Stream Recordset Copies Are Completely Separate Objects From Their Original Recordsets"<br>
<br>
 End If <br>
<br>
End Sub</font><br>
<br>
<br>
<font color="#0000FF">Function CreateFabricatedRecordset() As ADODB.Recordset<br>
 'Creates a Fabricated Recordser populated with names<br>
<br>
 Dim rst As ADODB.Recordset<br>
 Dim varField As Variant<br>
<br>
<br>
 'Create a Fabricated Recordset<br>
 Set rst = New ADODB.Recordset<br>
<br>
 With rst.Fields<br>
   .Append "LastName", adVarChar, 20<br>
 End With<br>
<br>
 'Open the recordset<br>
 rst.Open<br>
<br>
<br>
 varField = Array("LastName")<br>
<br>
<br>
 'Populate the fabricated Recordset with 3 names<br>
 With rst<br>
  .AddNew varField, Array("John")<br>
  .AddNew varField, Array("Paul")<br>
  .AddNew varField, Array("King")<br>
 End With<br>
<br>
<br>
 'Move to the first record<br>
 rst.MoveFirst<br>
<br>
<br>
 'Return the created Recordset<br>
 Set CreateFabricatedRecordset = rst <br>
<br>
End Function</font></font></p>
<p> </p>
<p><font face="Garamond"><b>ATTACHED FILES</b></font><font face="Garamond"><b>:<br>
</b>The Visual Basic implementation of this source code and a Microsoft Word(c)
2000 documentation of the project.<br>
<br>
<br>
<b>CONCLUSION:</b><br>
The ADO 2.5 Stream object can be used to create distinctly separate Recordset objects which the ADO Recordset.Clone method cannot do.<br>
<br>
<br>
<b>AUTHOR:</b><br>
Kingsley is a Technology Consultant specializing in Business Intelligence and can be reached at his site http://www.visual-basic-data-mining.net<br>
<br>
http://www.visual-basic-data-mining.net is a site dedicated to Visual Basic Data Mining Source Code distribution.Kingsley has made the Data Mining Source Code freely available to the public at http://www.visual-basic-data-mining.net<br>
</font></p>
<p><font face="Garamond"><br>
</font></p>
</body>
</html>

