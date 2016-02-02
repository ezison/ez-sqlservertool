' -------------------------------------------------------------------------------
'  DBStoredGet.vbs - DB�X�g�A�h�擾
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

' -----------------------------------
' ���ʃ��W���[����ǂݍ���
' -----------------------------------
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")

Dim pluginPath
pluginPath = Plugin.GetPluginDir()

Execute fs.OpenTextFile(fs.BuildPath(pluginPath, "vbs\Common.vbs")).ReadAll()

Call Main

Sub Main

	Dim fs
	Set fs = CreateObject("Scripting.FileSystemObject")

	Dim pluginPath
	pluginPath = Plugin.GetPluginDir()

	Dim selectedString
	selectedString = Editor.GetSelectedString

	If selectedString = "" Then
		SelectWord
		selectedString = Editor.GetSelectedString
	End If

	selectedString = Replace(selectedString, "[", "")
	selectedString = Replace(selectedString, "]", "")

	' -----------------------------------
	' INI�t�@�C������DB�ڑ������擾
	' -----------------------------------
	Dim ini
	Set ini = new IniFile
	ini.FilePath = fs.BuildPath(pluginPath, "exe\DBConnection.ini")

	Dim dbServer
	Dim dbDatabase
	Dim dbSspi
	Dim dbLoginId
	Dim dbPassword

	dbServer = ini.ReadString("DBConnectInfo_1", "Server", "")
	dbDatabase = ini.ReadString("DBConnectInfo_1", "Database", "")
	dbLoginId = ini.ReadString("DBConnectInfo_1", "LoginId", "")
	dbPassword = ini.ReadString("DBConnectInfo_1", "Password", "")

	If ini.ReadString("DBConnectInfo_1", "SSPI", "False") = "True" Then
		dbSspi = true
	Else
		dbSspi = false
	End If

	' -----------------------------------
	' DB�ڑ�����
	' -----------------------------------
	Dim connection
	Set connection = DBConnect(dbServer, dbDatabase, dbSspi, dbLoginId, dbPassword)

	' -----------------------------------
	' �X�g�A�h�擾����
	' -----------------------------------
	Dim storedSql
	storedSql = DBGetStored(connection, selectedString)

	' -----------------------------------
	' �X�g�A�h�ۑ����t�@�C���I�[�v��
	' -----------------------------------
	If storedSql <> "" Then

		Dim objTempFolder
		Set objTempFolder = fs.getSpecialFolder(2)

		Dim strTempFolder
		strTempFolder = fs.BuildPath(objTempFolder.Path, "dbstored\" & dbDatabase)

		If Not fs.FolderExists(strTempFolder) Then
			CreateFolder strTempFolder
		End If

		Dim strFilePath
		strFilePath = fs.BuildPath(strTempFolder, selectedString & ".sql")
		
		' �t�@�C�����r�����b�N���ł��邩�����݂�
		strFilePath = TryExclusiveOpenFile(strFilePath)

		Dim fsWriter
		Set fsWriter = fs.CreateTextFile(strFilePath, True) 
		fsWriter.Write (storedSql) 
		fsWriter.Close 
		
		Editor.FileOpen(strFilePath)

	Else
	
		MsgBox selectedString & "��������܂���B"
		
	End If
	
End Sub
