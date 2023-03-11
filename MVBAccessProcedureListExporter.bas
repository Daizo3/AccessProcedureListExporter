Attribute VB_Name = "MVBAccessProcedureListExporter"
' MVBAccessProcedureListExporter
' (2023/03/11 D) created
'
' Copyright 2023 D.Yamazaki, Ammres Corp.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation
' files (the �gSoftware�h), to deal in the Software without restriction, including without limitation the rights to use, copy, modify,
' merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED �gAS IS�h, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT
' LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH
' THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'
' (ja-jp)
' �ȉ��ɒ�߂�����ɏ]���A�{�\�t�g�E�F�A����ъ֘A�����̃t�@�C���i�ȉ��u�\�t�g�E�F�A�v�j�̕������擾���邷�ׂĂ̐l�ɑ΂��A
' �\�t�g�E�F�A�𖳐����Ɉ������Ƃ𖳏��ŋ����܂��B����ɂ́A�\�t�g�E�F�A�̕������g�p�A���ʁA�ύX�A�����A�f�ځA�Еz�A�T�u���C�Z���X�A
' �����/�܂��͔̔����錠���A����у\�t�g�E�F�A��񋟂��鑊��ɓ������Ƃ������錠�����������Ɋ܂܂�܂��B
'
' ��L�̒��쌠�\������і{�����\����\�t�g�E�F�A�̂��ׂĂ̕����܂��͏d�v�ȕ����ɋL�ڂ�����̂Ƃ��܂��
'
' �\�t�g�E�F�A�́u����̂܂܁v�ŁA�����ł��邩�Öقł��邩���킸�A����̕ۏ؂��Ȃ��񋟂���܂��B�����ł����ۏ؂Ƃ́A���i���A
' ����̖ړI�ւ̓K�����A����ь�����N�Q�ɂ��Ă̕ۏ؂��܂݂܂����A����Ɍ��肳�����̂ł͂���܂���B ��҂܂��͒��쌠�҂́A
' �_��s�ׁA�s�@�s�ׁA�܂��͂���ȊO�ł��낤�ƁA�\�t�g�E�F�A�ɋN���܂��͊֘A���A���邢�̓\�t�g�E�F�A�̎g�p�܂��͂��̑��̈����ɂ����
' �������؂̐����A���Q�A���̑��̋`���ɂ��ĉ���̐ӔC������Ȃ����̂Ƃ��܂��B
'
'The MIT License ? Open Source Initiative
'https://opensource.org/license/mit/
'
'https://licenses.opensource.jp/MIT/MIT.html
'https://licenses.opensource.jp/MIT/MIT.html
Option Compare Database
Option Explicit

' Export current Access VBA procedure list to csv file
' (ja-jp) ���݂� Access �v���W�F�N�g�i�t�@�C���j���� VBA �v���V�[�W���E���X�g�� csv �t�@�C���ɏo�͂���
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
