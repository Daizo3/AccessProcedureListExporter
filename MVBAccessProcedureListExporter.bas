Attribute VB_Name = "MVBAccessProcedureListExporter"
' MVBAccessProcedureListExporter
' (2023/03/11 D) created
'
' Copyright 2023 D.Yamazaki, Ammres Corp.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation
' files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify,
' merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT
' LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH
' THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'
' (ja-jp)
' 以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル（以下「ソフトウェア」）の複製を取得するすべての人に対し、
' ソフトウェアを無制限に扱うことを無償で許可します。これには、ソフトウェアの複製を使用、複写、変更、結合、掲載、頒布、サブライセンス、
' および/または販売する権利、およびソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。
'
' 上記の著作権表示および本許諾表示を､ソフトウェアのすべての複製または重要な部分に記載するものとします｡
'
' ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、何らの保証もなく提供されます。ここでいう保証とは、商品性、
' 特定の目的への適合性、および権利非侵害についての保証も含みますが、それに限定されるものではありません。 作者または著作権者は、
' 契約行為、不法行為、またはそれ以外であろうと、ソフトウェアに起因または関連し、あるいはソフトウェアの使用またはその他の扱いによって
' 生じる一切の請求、損害、その他の義務について何らの責任も負わないものとします。
'
'The MIT License ? Open Source Initiative
'https://opensource.org/license/mit/
'
'https://licenses.opensource.jp/MIT/MIT.html
'https://licenses.opensource.jp/MIT/MIT.html
Option Compare Database
Option Explicit

' Export current Access VBA procedure list to csv file
' (ja-jp) 現在の Access プロジェクト（ファイル）中の VBA プロシージャ・リストを csv ファイルに出力する
' Licence: MIT
' (2023/03/11 D) created
Sub AccessProcedureListExporter()
    
    Const filename As String = "procedure_list.csv"     ' Change by your oppotunity
    
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    Dim folder As String
    folder = wsh.SpecialFolders("MyDocuments")
    
    Set wsh = Nothing
    
    Dim project As String
    project = Replace(CurrentProject.Name, ".", "_")
    
    Dim fullpath As String
    fullpath = folder & "\" & project & "_" & filename
    
    Open fullpath For Output As #1
    
    Const csvHeader As String = """Module Number"",""Module Name"",""Procedure Number"",""Procedure Name"",""At Line"""
    Print #1, csvHeader
    
    Dim moduleCount As Integer
    moduleCount = Application.VBE.ActiveVBProject.VBComponents.Count
    
    Dim moduleDim As Integer
    For moduleDim = 1 To moduleCount
        Dim moduleName As String
        moduleName = Application.VBE.ActiveVBProject.VBComponents(moduleDim).Name
        
        With Application.VBE.ActiveVBProject.VBComponents(moduleDim).CodeModule
            
            Dim procLine As Integer
            Dim procName As String
            Dim procCount As Integer
            procCount = 0
            For procLine = 1 To .CountOfLines
                If procName <> .ProcOfLine(procLine, 0) Then
                    procName = .ProcOfLine(procLine, 0)
                    If (Len(procName) > 0) Then
                        procCount = procCount + 1
                        Dim csvLine As String
                        csvLine = moduleDim & ",""" & moduleName & """," & procCount & ",""" & procName & """, " & procLine
                        Print #1, csvLine
                        Debug.Print csvLine
                    End If
                End If
            Next procLine
        End With
        
    Next moduleDim
    
    Close #1
    
    MsgBox "Exported procedure list : " & fullpath, vbOKOnly + vbInformation

End Sub
