' -------------------------------------------------------------------------------
'  DBConnection.vbs - DB�ڑ�
' -------------------------------------------------------------------------------
' 
'  Copyright(c) 2016 EZOLAB Co., Ltd. All Rights Reserved.
' 
'  The MIT License
' 
'  Permission is hereby granted, free of charge, to any person obtaining a copy
'  of this software and associated documentation files (the "Software"), to deal
'  in the Software without restriction, including without limitation the rights
'  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'  copies of the Software, and to permit persons to whom the Software is
'  furnished to do so, subject to the following conditions:
' 
'  The above copyright notice and this permission notice shall be included in
'  all copies or substantial portions of the Software.
' 
'  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
'  THE SOFTWARE.
' 
' -------------------------------------------------------------------------------
Option Explicit

' ==============================================================================
' Summary : DB�ɐڑ�����B
' Inputs  : server    : �T�[�o�[
'           database  : �f�[�^�x�[�X
'           sspi      : SSPI
'           userId    : ���[�U�[ID
'           password  : �p�X���[�h
' Returns : DB�R�l�N�V����
' ==============================================================================
Function DBConnect(ByVal server, ByVal database, ByVal sspi, ByVal userId, ByVal password)

	Dim connection

	Set connection = createobject("ADODB.Connection") 
	
	If sspi = True Then
		connection.connectionstring="provider=sqloledb" & "; Data Source=" & server & "; Integrated Security=SSPI; Initial Catalog=" & database
	Else
		connection.connectionstring="provider=sqloledb" & "; Data Source=" & server & "; User ID=" & userId & "; Password=" & password & "; Initial Catalog=" & database
	End If

	connection.open
	' �N���C�A���g�T�C�h�J�[�\�����g�p����
	connection.CursorLocation = 3
	
	Set DBConnect = connection

End Function

' ==============================================================================
' Summary : DB����ؒf����B
' Inputs  : connection    : DB�R�l�N�V����
' ==============================================================================
Sub DBClose(ByVal connection)

	If Not connection Is Nothing Then
		connection.Close
		Set connection = Nothing
	End If

End Sub

' ==============================================================================
' Summary : SELECT�����s����B
' Inputs  : connection    : DB�R�l�N�V����
'           sql           : SQL
' Returns : ���R�[�h���X�g
' ==============================================================================
Function DBExecSelect(ByVal connection, ByVal sql)

	Dim ret()

	Dim rst 
	Dim recordCount
	Dim i, j
	
	Dim schema
	Dim objName
	
	Set rst = connection.execute(sql)
	
	' ���R�[�h�����擾����
	recordCount = rst.RecordCount
	
	' ���R�[�h��0���̏ꍇ
	If recordCount = 0 Then

		DBExecSelect = Empty
		Exit Function
	End If
	
	' �߂�l���m�ۂ���
	Redim ret(recordCount, rst.Fields.Count)
	
	i = 0
	Do Until rst.eof 

		For j = 0 To rst.fields.count - 1
			ret(i, j) = rst.Fields(j).Value
		Next 

		rst.MoveNext
		
		i = i + 1
	Loop
	
	If Not rst Is Nothing Then
		rst.close
		Set rst = Nothing
	End If

	DBExecSelect = ret

End Function

' ==============================================================================
' Summary : �X�g�A�h���e���擾����B
' Inputs  : connection    : DB�R�l�N�V����
'           storedName    : �X�g�A�h��
' Returns : �X�g�A�h���e
' ==============================================================================
Function DBGetStored(ByVal connection, ByVal storedName)

	DBGetStored = ""

	Dim i
	Dim ret
	
	Dim schema
	Dim objName
	
	schema = "dbo"
	objName = storedName

	ret = DBExecSelect(connection, "SELECT m.definition FROM sys.sql_modules AS m INNER JOIN sys.objects AS o ON m.object_id = o.object_id WHERE SCHEMA_NAME(o.schema_id) = '" & schema & "' AND OBJECT_NAME(o.object_id) = '" & objName & "'")
	
	If Not IsArray(ret) Then
		Exit Function
	End If

	For i = LBound(ret) To UBound(ret) - 1
		DBGetStored = ret(i, 0)
	Next
	
End Function

Dim arr, i
arr = DBParseObject("[[[a 1]]]]].[b ]]]]2].[]]]]]]]]1 c 3eeeeeeeee]]")
arr = DBParseObject("...a.b����������.].[]]ab]]cdefg]]]")

For i = 0 To Ubound(arr) - 1
	MsgBox arr(i)
Next

' ==============================================================================
' Summary : DB�̃I�u�W�F�N�g������͂���B
' Inputs  : objName    : �I�u�W�F�N�g��
' Returns : �I�u�W�F�N�g������͂����z��
' ==============================================================================
Function DBParseObject(ByVal objName)

	' �߂�l
	Dim ret()
	
	' �߂�l�z��̒���
	Dim retLen
	retLen = 0
	
	' �߂�l�z��̌��݂̃C���f�b�N�X
	Dim retIndex
	retIndex = 0

	' �����J�n�C���f�b�N�X
	Dim ib
	
	' �����C���f�b�N�X
	Dim i
	
	' �L�����N�^����
	Dim c
	' ���̃L�����N�^����
	Dim cNext
	cNext = ""
	
	' �p���ʂ̊J�n�ʒu�i0�̏ꍇ�́A�J�n���Ă��Ȃ��j
	Dim beginBracket
	beginBracket = 0
	
	Dim endBracket

	ib = 1
	For i = 1 To Len(objName)
	
		' ���݂̕������擾
		c = Mid(objName, i, 1)
		' ���̕������擾
		cNext = Mid(objName, i + 1, 1)
		
		If c = "[" And beginBracket = 0 Then
			' �p���ʂ̊J�n�C���f�b�N�X
		
			beginBracket = i
			ib = i + 1
			
		ElseIf c = "]" And cNext = "]" And beginBracket <> 0 Then
		
			i = i + 1
		
		ElseIf c = "]" And cNext <> "]" And beginBracket <> 0 Then
		
			If i - ib > 0 Then
			
				retLen = retLen + 1
				Redim Preserve ret(retLen)
				
				ret(retIndex) = Replace(Mid(objName, ib, i - ib), "]]", "]")
				retIndex = retIndex + 1
				
			End If
			
			' �����J�n�ʒu���Đݒ肷��
			ib = i + 1
			
			' �p���ʂ̏I��
			beginBracket = 0
		
		ElseIf c = "." Then
			' �h�b�g���o�������̂ŁA��؂��\���Ɣ��f����
		
			If beginBracket = 0 Then
				' �p���ʂ��J�n���Ă��Ȃ��ꍇ�́A�h�b�g�ȑO�̕�������擾����
			
				If i - ib > 0 Then
			
					' �z��̃T�C�Y���g������
					retLen = retLen + 1
					Redim Preserve ret(retLen)
					
					' �z��ɃI�u�W�F�N�g�����i�[����
					ret(retIndex) = Replace(Mid(objName, ib, i - ib), "]]", "]")
					retIndex = retIndex + 1
					
				End If
				
				' �����J�n�ʒu���Đݒ肷��
				ib = i + 1
			
			End IF
		
		End If
		
	Next

	If i - ib > 0 Then
	
		retLen = retLen + 1
		Redim Preserve ret(retLen)
		
		ret(retIndex) = Replace(Mid(objName, ib, i - ib), "]]", "]")
		retIndex = retIndex + 1
		
	End If

	If retLen = 0 Then
		Redim ret(1)
		ret(0) = objName
	End If

	DBParseObject = ret

End Function

' ==============================================================================
' Summary : ���̃v���Z�X�ŊJ����Ă��Ȃ������`�F�b�N����B
' Inputs  : fileName    : �t�@�C����
' Returns : True �r�����b�N���AFalse �ʏ�
' ==============================================================================
Function IsExclusiveLockFile(ByVal fileName)

	Dim fs, f
	Set fs = CreateObject("Scripting.FileSystemObject")

	If Not fs.FileExists(fileName) Then
		' �t�@�C�������݂��Ȃ��̂ŏI��
		Exit Function
	End If
	
	On Error Resume Next
	
	' �������ݐ�p�ŃI�[�v���ŃG���[���������Ȃ���΁A�N���J���Ă��Ȃ��Ƃ݂Ȃ�
	Set f = fs.OpenTextFile(fileName, 2, false)
	
	If Err.Number = 0 Then
		IsExclusiveLockFile = False
	Else
		IsExclusiveLockFile = True
		Exit Function
	End If

	Err.Clear
	On Error Goto 0

	If Not f Is Nothing Then
		f.Close
	End If
	
	Set fs = Nothing
	
End Function

' ==============================================================================
' Summary : �t�@�C���̍폜�����݂�B�t�@�C���̍폜�Ɏ��s�����ꍇ��
'           �t�@�C�����̖����ɃC���f�b�N�X��t�^���A���݂��Ȃ��t�@�C������ԋp����B
' Inputs  : fileName    : �t�@�C���p�X
' Returns : �폜�ɐ��������t�@�C���p�X
' ==============================================================================
Function TryExclusiveOpenFile(ByVal fileName)

	Dim fs
	Set fs = CreateObject("Scripting.FileSystemObject")

	Dim folder
	folder = fs.GetParentFolderName(fileName)
	
	Dim fileNameWithoutExt
	fileNameWithoutExt = fs.GetBaseName(fileName)

	Dim fileNameOnlyExt
	fileNameOnlyExt = fs.GetExtensionName(fileName)
	
	Dim deleteFilePath
	deleteFilePath = fileName
	
	Dim i
	i = 0
	
	Do While True
	
		If Not IsExclusiveLockFile(deleteFilePath) Then
			' �t�@�C�������݂��Ȃ��̂ŏI��
			TryExclusiveOpenFile = deleteFilePath
			Exit Do
		End If
		
		' �t�@�C�����𐶐�����
		i = i + 1
		deleteFilePath = fs.BuildPath(folder, fileNameWithoutExt & "_" & i & "." & fileNameOnlyExt)
	Loop
	
	Set fs = Nothing

End Function

' ==============================================================================
' Summary : �t�H���_�𐶐�����B�T�u�t�H���_���܂߂č쐬����B
' Inputs  : fileName    : �t�@�C���p�X
' Returns : �폜�ɐ��������t�@�C���p�X
' ==============================================================================
Function CreateFolder(ByVal fileName)

	Dim oShell
	Set oShell = CreateObject("WSCript.Shell")
	
	Dim ret
	
	ret = oShell.run ("cmd /c MD " & fileName, 0, 1)
	
	If ret <> 0 Then
		
		Err.Raise 51, "Common.vbs", "Folder create failed. path = " & fileName
	
	End If
	
	Set oShell = Nothing

End Function

' ==============================================================================
' Summary : Ini�t�@�C������N���X�B
' ==============================================================================
Class IniFile

	Private fso_
	Private f_

	Private iniFilePath_ ' Path to the ini File
	Private section_     ' [section]
	Private key_         ' Key=Value
	Private default_     ' Return it when an error occurs
	Private content_

	' ==============================================================================
	' Summary : �R���X�g���N�^�B
	' ==============================================================================
	Private Sub Class_Initialize
		default_ = ""
		Set fso_ = CreateObject("Scripting.FileSystemObject")
	End Sub
	
	' ==============================================================================
	' Summary : �f�X�g���N�^�B
	' ==============================================================================
	Private Sub Class_Terminate
		Call Save
		Set fso_ = Nothing
	End Sub
	
	' ==============================================================================
	' Summary : �t�@�C���p�X�v���p�e�B�B
	' Inputs  : FileName  : �t�@�C����
	' ==============================================================================
	Property Let FilePath(ByVal FileName)
	
		iniFilePath_ = FileName
		
		If fso_.FileExists(iniFilePath_) Then
		
			Set f_ = fso_.OpenTextFile(iniFilePath_, 1)
			
			If f_.AtEndOfStream = False Then
				content_ = f_.ReadAll
			Else
				content_ = ""
			End If
			
			f_.close
			Set f_ = Nothing
			
		Else
			content_ = ""
			
		End If
		
	End Property

	' ==============================================================================
	' Summary : �ۑ�����B
	' Inputs  : 
	' ==============================================================================
	Public Sub Save()

		' Create a brand new ini file
		
		Set f_ = fso_.CreateTextFile(iniFilePath_, True)
		f_.Write content_
		f_.Close
		
		Set f_ = Nothing
		
	End Sub

	' ==============================================================================
	' Summary : �R���e���c���X�g���擾����B
	' Inputs  : 
	' Returns : �R���e���c���X�g
	' ==============================================================================
	Private Property Get ContentArray()
	
		' All the file in an array of lines
		ContentArray = Split(content_, vbCrLf, -1, 1)
	End Property

	' ==============================================================================
	' Summary : �Z�N�V��������������B
	' Inputs  : StartLine : �J�n�s
	'           EndLine   : �I���s
	' ==============================================================================
	Private Sub FindSection(ByRef StartLine, ByRef EndLine)
	
		Dim x, A, s
		
		StartLine = -1
		EndLine   = -2
		
		A = ContentArray
		
		For x = 0 To UBound(A)
		
			s = UCase(Trim(A(x)))
			If s = "[" & UCase(section_) & "]" Then
				StartLine = x
			Else
			
				If (Left(s,1) = "[") And (Right(s,1) = "]") Then
				
					If StartLine >= 0 Then
					
						EndLine = x - 1
						
						' A Space before the next section ?
						If EndLine>0 Then
							If Trim(A(EndLine)) = "" Then EndLine = EndLine-1
						End If
						
						Exit Sub
						
					End If
					
				End If
				
			End If
		Next
		
		If (StartLine > 0) And (EndLine < 0) Then EndLine  =  UBound(A)
		
	End Sub

	' ==============================================================================
	' Summary : �l���擾����B
	' Inputs  : 
	' Returns : �l
	' ==============================================================================
	Private Property Get Value()
	
		' Retrieve the value for the current key in the current section
		Dim x, i, j, A, s
		
		FindSection i, j
		
		A = ContentArray
		Value = default_
		
		' Search only in the good section
		For x = i + 1 To j
		
			s = Trim(A(x))
			
			If UCase(Left(s, Len(key_))) = UCase(key_) Then
			
				Select Case Mid(s, Len(key_) + 1, 1)
				
				Case "="
				
					Value = Trim(Mid(s, Len(key_) + 2))
					Exit Property
					
				Case " ", chr(9)
				
					x = Instr(Len(key_), s, "=")
					Value = Trim(Mid(s, x + 1))
					Exit Property
					
				End Select
				
			End If
		Next
		
	End Property
	
	' ==============================================================================
	' Summary : �l��ݒ肷��B
	' Inputs  : sValue : �l
	' ==============================================================================
	Private Property Let Value(sValue)
	
		' Write the value for a key in a section
		Dim i, j, A, x, s, f
		
		FindSection i, j
		
		If i < 0 Then ' Session doesn't exist
			content_ = content_ & vbCrLf & "[" & section_ & "]" & vbCrLf & key_ & "=" & sValue
		Else
		
			A = ContentArray
			f = -1
			
			'Search for the key, either the key exists or not
			For x = i + 1 To j
			
				s = Trim(A(x))
				If UCase(Left(s,Len(key_))) = UCase(key_) Then
				
					Select Case Mid(s, Len(key_) + 1, 1)
					Case " ", chr(9), "="
						f = x 'Key found
						A(x) = key_ & "=" & sValue
					End Select
					
				End If
				
			Next
			
			If f = -1 Then
			
				' Not found, add it at the end of the section
				Redim Preserve A(UBound(A) + 1)
				
				For x = UBound(A) To j + 2 Step -1
					A(x)=A(x-1)
				Next
				
				A(j + 1) = key_ & "=" & sValue
				
			End If
			
			' Define the content
			s = ""
			
			For x = 0 To UBound(A)	
				s = s & A(x) & vbCrLf
			Next
			
			' Suppress the last CRLF
			If Right(s, 2) = vbCrLf Then s = Left(s, Len(s) - 2)
			
			content_ = s ' Write it
			
		End If
		
	End Property
	
	' ==============================================================================
	' Summary : ��������������ށB
	' Inputs  : sSection : �Z�N�V����
	'         : sKey     : �L�[
	'         : sValue   : �l
	' ==============================================================================
	Public Sub WriteString(sSection, sKey, sValue)
	
		section_ = sSection
		key_ = sKey
		Value = sValue
		
	End Sub
	
	' ==============================================================================
	' Summary : �l���擾����B
	' Inputs  : sSection   : �Z�N�V����
	'         : sKey       : �L�[
	'         : sDefault   : �f�t�H���g
	' Returns : �l
	' ==============================================================================
	Public Function ReadString(sSection, sKey, sDefault)
	
		section_ = sSection
		key_ = sKey
		default_ = sDefault
		
		ReadString = Value
		
	End Function
	
End Class
