' -------------------------------------------------------------------------------
'  DBConnection.vbs - DB接続
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
' Summary : DBに接続する。
' Inputs  : server    : サーバー
'           database  : データベース
'           sspi      : SSPI
'           userId    : ユーザーID
'           password  : パスワード
' Returns : DBコネクション
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
	' クライアントサイドカーソルを使用する
	connection.CursorLocation = 3
	
	Set DBConnect = connection

End Function

' ==============================================================================
' Summary : DBから切断する。
' Inputs  : connection    : DBコネクション
' ==============================================================================
Sub DBClose(ByVal connection)

	If Not connection Is Nothing Then
		connection.Close
		Set connection = Nothing
	End If

End Sub

' ==============================================================================
' Summary : SELECTを実行する。
' Inputs  : connection    : DBコネクション
'           sql           : SQL
' Returns : レコードリスト
' ==============================================================================
Function DBExecSelect(ByVal connection, ByVal sql)

	Dim ret()

	Dim rst 
	Dim recordCount
	Dim i, j
	
	Dim schema
	Dim objName
	
	Set rst = connection.execute(sql)
	
	' レコード数を取得する
	recordCount = rst.RecordCount
	
	' レコードが0件の場合
	If recordCount = 0 Then

		DBExecSelect = Empty
		Exit Function
	End If
	
	' 戻り値を確保する
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
' Summary : ストアド内容を取得する。
' Inputs  : connection    : DBコネクション
'           storedName    : ストアド名
' Returns : ストアド内容
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
arr = DBParseObject("...a.bあいうえお.].[]]ab]]cdefg]]]")

For i = 0 To Ubound(arr) - 1
	MsgBox arr(i)
Next

' ==============================================================================
' Summary : DBのオブジェクト名を解析する。
' Inputs  : objName    : オブジェクト名
' Returns : オブジェクト名を解析した配列
' ==============================================================================
Function DBParseObject(ByVal objName)

	' 戻り値
	Dim ret()
	
	' 戻り値配列の長さ
	Dim retLen
	retLen = 0
	
	' 戻り値配列の現在のインデックス
	Dim retIndex
	retIndex = 0

	' 文字開始インデックス
	Dim ib
	
	' 文字インデックス
	Dim i
	
	' キャラクタ文字
	Dim c
	' 次のキャラクタ文字
	Dim cNext
	cNext = ""
	
	' 角括弧の開始位置（0の場合は、開始していない）
	Dim beginBracket
	beginBracket = 0
	
	Dim endBracket

	ib = 1
	For i = 1 To Len(objName)
	
		' 現在の文字を取得
		c = Mid(objName, i, 1)
		' 次の文字を取得
		cNext = Mid(objName, i + 1, 1)
		
		If c = "[" And beginBracket = 0 Then
			' 角括弧の開始インデックス
		
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
			
			' 文字開始位置を再設定する
			ib = i + 1
			
			' 角括弧の終了
			beginBracket = 0
		
		ElseIf c = "." Then
			' ドットが出現したので、区切りを表すと判断する
		
			If beginBracket = 0 Then
				' 角括弧が開始していない場合は、ドット以前の文字列を取得する
			
				If i - ib > 0 Then
			
					' 配列のサイズを拡張する
					retLen = retLen + 1
					Redim Preserve ret(retLen)
					
					' 配列にオブジェクト名を格納する
					ret(retIndex) = Replace(Mid(objName, ib, i - ib), "]]", "]")
					retIndex = retIndex + 1
					
				End If
				
				' 文字開始位置を再設定する
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
' Summary : 他のプロセスで開かれていないかをチェックする。
' Inputs  : fileName    : ファイル名
' Returns : True 排他ロック中、False 通常
' ==============================================================================
Function IsExclusiveLockFile(ByVal fileName)

	Dim fs, f
	Set fs = CreateObject("Scripting.FileSystemObject")

	If Not fs.FileExists(fileName) Then
		' ファイルが存在しないので終了
		Exit Function
	End If
	
	On Error Resume Next
	
	' 書き込み専用でオープンでエラーが発生しなければ、誰も開いていないとみなす
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
' Summary : ファイルの削除を試みる。ファイルの削除に失敗した場合は
'           ファイル名の末尾にインデックスを付与し、存在しないファイル名を返却する。
' Inputs  : fileName    : ファイルパス
' Returns : 削除に成功したファイルパス
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
			' ファイルが存在しないので終了
			TryExclusiveOpenFile = deleteFilePath
			Exit Do
		End If
		
		' ファイル名を生成する
		i = i + 1
		deleteFilePath = fs.BuildPath(folder, fileNameWithoutExt & "_" & i & "." & fileNameOnlyExt)
	Loop
	
	Set fs = Nothing

End Function

' ==============================================================================
' Summary : フォルダを生成する。サブフォルダも含めて作成する。
' Inputs  : fileName    : ファイルパス
' Returns : 削除に成功したファイルパス
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
' Summary : Iniファイル操作クラス。
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
	' Summary : コンストラクタ。
	' ==============================================================================
	Private Sub Class_Initialize
		default_ = ""
		Set fso_ = CreateObject("Scripting.FileSystemObject")
	End Sub
	
	' ==============================================================================
	' Summary : デストラクタ。
	' ==============================================================================
	Private Sub Class_Terminate
		Call Save
		Set fso_ = Nothing
	End Sub
	
	' ==============================================================================
	' Summary : ファイルパスプロパティ。
	' Inputs  : FileName  : ファイル名
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
	' Summary : 保存する。
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
	' Summary : コンテンツリストを取得する。
	' Inputs  : 
	' Returns : コンテンツリスト
	' ==============================================================================
	Private Property Get ContentArray()
	
		' All the file in an array of lines
		ContentArray = Split(content_, vbCrLf, -1, 1)
	End Property

	' ==============================================================================
	' Summary : セクションを検索する。
	' Inputs  : StartLine : 開始行
	'           EndLine   : 終了行
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
	' Summary : 値を取得する。
	' Inputs  : 
	' Returns : 値
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
	' Summary : 値を設定する。
	' Inputs  : sValue : 値
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
	' Summary : 文字列を書き込む。
	' Inputs  : sSection : セクション
	'         : sKey     : キー
	'         : sValue   : 値
	' ==============================================================================
	Public Sub WriteString(sSection, sKey, sValue)
	
		section_ = sSection
		key_ = sKey
		Value = sValue
		
	End Sub
	
	' ==============================================================================
	' Summary : 値を取得する。
	' Inputs  : sSection   : セクション
	'         : sKey       : キー
	'         : sDefault   : デフォルト
	' Returns : 値
	' ==============================================================================
	Public Function ReadString(sSection, sKey, sDefault)
	
		section_ = sSection
		key_ = sKey
		default_ = sDefault
		
		ReadString = Value
		
	End Function
	
End Class
