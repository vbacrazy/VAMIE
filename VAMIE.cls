VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBA_VAMIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' VAMIE (VBA Auto Mation for Internet Explorer)
'
' LastModified:
'   2014/9/2    renewal
'
'   2014/1/8    + DOM�Z���N�^�g�� (ByName��ǉ�)
'
'   2014/1/7    * Vbaget�V���[�Y�Ƃ��ă��t�@�C��(���U���Ă����R�[�h���W��)
'                   + �v���p�e�B�̐���
'                   + ����R�}���h�̒ǉ�(goto_url_NoWait, Wait, DisableConfirmFunction)
'                   + 64bit�ւ̑Ή�
'
'   2012/6/20   - domselec()�̃G���[�����폜
'   2012/6/18   + domselec()�ɃG���[�����ǉ� �i�f�o�b�O���ɖ���c�����Â炢���߁j
'
' This Class Module:
'   First Release 2011/11/08
'   Created By D*isuke YAMAKWA
'
' SpecialThanks:
'   Excel VBA�̃}�N����, IE���������삵�悤(DOM�Z���N�^�֐���VBA�Ŏ���)
'   http://d.hatena.ne.jp/language_and_engineering/20090710/p1
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private ie As Object
Dim flagQuitWhenTerminate As Boolean ' �f�X�g���N�^�p

Public Property Let Visible(setBoolean)
    ie.Visible = setBoolean
End Property
Public Property Get Visible()
    Visible = ie.Visible
End Property
Public Property Let FullScreen(setBoolean)
    ie.FullScreen = setBoolean
End Property
Public Property Get FullScreen()
    FullScreen = ie.FullScreen
End Property
Public Property Get Document() 'VAMIE�ɗp�ӂ��ꂽ���\�b�h�ł͖ړI�̓��삪�o���Ȃ����p(Document�N���X�𒼐ڑ��삵�����ꍇ�p)
    Document = ie.Document
End Property
Public Property Get LocationURL()
        LocationURL = ie.LocationURL
End Property
Public Property Get LocationName()
        LocationName = ie.LocationName
End Property

Sub Class_Initialize()
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
End Sub
Sub Class_Terminate()
        If flagQuitWhenTerminate Then ie.Quit
        Set ie = Nothing
End Sub

Public Property Let AutoQuit(setBoolean)
        flagQuitWhenTerminate = setBoolean
End Property

'--------------------   '--------------------   '--------------------
Public Sub Navigate(url)
    ie.Navigate url
    WaitLoading
End Sub
Public Sub NavigateWithNoWait(url)      ' WaitLoading�����ނƖ������[�v����悤�ȃy�[�W�΍�
    ie.Navigate url
End Sub
Public Sub Quit()
    ie.Quit
End Sub

Sub ResizeTo(width, height)
        If LocationURL = Empty Then
                msgbox ("VAMIE ���� : ReizeTo���\�b�h�̓y�[�W��\��������ŌĂяo���Ă�������")
                Exit Sub
        End If

        Call ExecuteJavaScript("window.resizeTo(" & width & "," & height & ");")
End Sub

'DOM�v�f����p���\�b�h�Q ----------------------------------------------------
Sub Exists(element)
        Dim test: Set test = element
        If test <> Empty Then
                Exists = True
        Else
                Exists = False
        End If
End Sub
Function Find(arr) ' �Ȉ�DOM�Z���N�^ �y�����̗^�����z��F VAMIE.Find(Array("id","hoge","class","fuga",1, "tag","table",2))
        Dim parent_obj: Set parent_obj = ie.Document
        Dim child_obj
        Dim dom_id, tag_name, index_num, name_

        Dim cur: cur = 0
        Dim continue_flag: continue_flag = True
        Do While continue_flag = True
                Select Case arr(cur):
                    Case "id"
                        dom_id = arr(cur + 1)
                        Set child_obj = parent_obj.getElementById(dom_id)
                        cur = cur + 2
                    Case "tag"
                        tag_name = arr(cur + 1)
                        index_num = arr(cur + 2)
                        Set child_obj = parent_obj.GetElementsByTagName(tag_name)(index_num)
                        cur = cur + 3
                    Case "name"
                        name_ = arr(cur + 1)
                        index_num = arr(cur + 2)
                        Set child_obj = parent_obj.GetElementsByName(name_)(index_num)
                        cur = cur + 3
                    Case "class"
                        name_ = arr(cur + 1)
                        index_num = arr(cur + 2)
                        Set child_obj = parent_obj.GetElementsByClassName(name_)(index_num)
                        cur = cur + 3
                End Select
                
                Set parent_obj = child_obj
                
                If cur > UBound(arr) Then
                    continue_flag = False
                End If
        Loop
        
        Set Find = parent_obj
End Function

Function FindById(dom_id)
    Set FindById = ie.Document.getElementById(dom_id) ' ���F��IE��getElementById��name���Q�Ƃ���
End Function
Function FindsByName(name)
    Set FindsByName = ie.Document.GetElementsByName(name)
End Function
Function FindsByTag(tag_name)
    Set FindsByTag = ie.Document.GetElementsByTagName(tag_name)
End Function
Function FindsByClass(className)
    Set FindsByClass = ie.Document.GetElementsByClassName(className)
End Function

Function GetInnerText(element) '�e�L�X�g���擾
    GetInnerText = element.innerText
End Function
Function GetInnerHTML(element) 'HTML�R�[�h���擾
    GetInnerHTML = element.innerHTML
End Function

Sub SetValue(element, val) ' �e�L�X�g�{�b�N�X�ւ̓��͂Ȃ�
    element.value = val
    WaitLoading
End Sub
Sub Click(element) ' ���M�{�^���⃊���N���N���b�N
    element.Click
    WaitLoading
End Sub
Sub SetCheckBox(element, checked_flag) ' �`�F�b�N�{�b�N�X�̏�Ԃ��Z�b�g���܂�
    If Not (element.Checked = checked_flag) Then
        Call Click(element)
    End If
End Sub
Sub SelectListBox(element, label) ' �Z���N�g�{�b�N�X�𕶌��x�[�X�őI�����܂�
    If Len(label) < 1 Then Exit Sub

    Dim opts: Set opts = element.Options
    Dim i: For i = 0 To opts.Length - 1
        If opts(i).innerText = label Then
            opts(i).Selected = True
            Exit Sub
        End If
    Next
End Sub
Sub SetRadioButton(element, value) ' ���W�I�{�^����l�x�[�X�őI�����܂�
    If Len(value) < 1 Then Exit Sub

    Dim radios: Set radios = element
    Dim i: For i = 0 To radios.Length - 1
        If radios(i).value = CStr(value) Then
            radios(i).Click
            Sleep 100
        End If
    Next
End Sub

' -----------------------------------------------------------------------------
Public Sub WaitLoading()
    Do While ie.Busy = True Or ie.ReadyState <> 4
        Sleep 100
        DoEvents
    Loop
    Sleep 100
End Sub
Public Sub Wait(millisecond)
    Sleep millisecond
    DoEvents
End Sub

' ���܂� ---------------------------------------------------------------------
Function GetIEVersion()
    Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
    Dim hoge: hoge = Fix(val(FS.GetFileVersion(ie.FullName)))
    GetIEVersion = hoge
End Function

Sub DisableConfirmFunction() 'confirm()�Ăяo�����Ɋm�F�_�C�A���O��\�������Ȃ�
    Dim ele: Set ele = ie.Document.createElement("SCRIPT")
    
    ele.Type = "text/javascript"
    ele.text = "function confirm() { return true; }"
    
    Call ie.Document.body.appendChild(ele)
End Sub

Sub Activate() 'SendKeys�p
        Dim wLoc, wSvc, wEnu, wIns
        Set wLoc = CreateObject("WbemScripting.SWbemLocator")
        Set wSvc = wLoc.ConnectServer
        Set wEnu = wSvc.InstancesOf("Win32_Process")
        Dim pId
        For Each wIns In wEnu
            If Not IsEmpty(wIns.ProcessId) And wIns.Description = "iexplore.exe" Then
                pId = wIns.ProcessId
            End If
        Next

        Dim wsh: Set wsh = CreateObject("Wscript.Shell")
        While Not wsh.AppActivate(pId)
                Sleep 100
        Wend
End Sub

Sub SendKeys(keys) '�l��������̂ł͂Ȃ��A�L�[���͂��G�~�����[�g�������ꍇ
        Dim wsh: Set wsh = CreateObject("Wscript.Shell")
        wsh.SendKeys keys
End Sub

Public Sub ExecuteJavaScript(jsCode)
        Call ie.Document.Script.setTimeout("javascript:" & jsCode, 1) ' ��2����:���s�܂ł̑ҋ@����[msec]
        WaitLoading
End Sub
