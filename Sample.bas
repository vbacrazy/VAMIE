Attribute VB_Name = "Sample"
Option Explicit

Sub SampleCodeAsReference()
        Dim ie: Set ie = New VAMIE

        ie.Visible = True '�f�t�H���gTrue
        'ie.Document 'Document�N���X�𒼐ڑ��삵�����Ƃ��p�̃v���p�e�B(Frame�y�[�W�̑���Ƃ�)
        ie.AutoQuit = True '�C���X�^���X�j������IE�����@�f�t�H���gFalse

        ie.FullScreen = True
        Call ie.ResizeTo(200, 300) '�E�B���h�E���T�C�Y

        ie.Activate '�E�B���h�E���A�N�e�B�u��
        Call ie.SendKeys("test") '�A�N�e�B�u�E�B���h�E��SendKeys ...�̂͂�

        Call ie.Navigate("http://www.google.com/") '�y�[�W��\�����ēǂݍ��݂��I���܂őҋ@
        Call ie.NavigateWithNoWait("http://www.google.com/") '�y�[�W��\���@�i�ҋ@�Ȃ��j�@���ҋ@��Ԃ���ɑ����y�[�W�΍�
        
        Dim nonsence
        nonsence = ie.LocationURL '���݂�URL���擾
        nonsence = ie.LocationName '���݂̃y�[�W�̃^�C�g�����擾

        Dim DOM_Element
        Set DOM_Element = ie.FindById("id") 'id���w�肵��DOM_Element���擾
        Set DOM_Element = ie.FindsByName("name")(0) 'Find[s]��DOM_Element�̔z���Ԃ��A�̂�DOM_Element���擾�������Ƃ��͓Y������
        Set DOM_Element = ie.FindsByClass("class")(0)
        Set DOM_Element = ie.FindsByTag("tag")(0)
        Set DOM_Element = ie.Find(Array("id", "res", "tag", "li", 0, "tag", "h3", 0))(0) 'DOM�Z���N�^�I�Ȃ�B�g����L�[���[�h��,id, name, tag, class

        If ie.Exists(DOM_Element) Then
                Call ie.GetInnerText(DOM_Element) '�e�L�X�g�擾
                Call ie.GetInnerHTML(DOM_Element) 'HTML�R�[�h�擾

                Call ie.SetValue(DOM_Element) '�l����͂Ă���������i�L�[���͂̃G�~�����[�V������SendKeys()�Łj
                Call ie.Click(DOM_Element) '�N���b�N�Ƃ�Submit
                Call ie.SetCheckBox(DOM_Element, True) '�`�F�b�N�{�b�N�X��ON/OFF�ݒ�
                Call ie.SelectListBox(DOM_Element, "label��") '���X�g�{�b�N�X�ɂ����āAlabel���ƈ�v����A�C�e����I��
                Call ie.SetRadioButton(DOM_Element, 3) '���W�I�{�^����l�x�[�X�őI��

                Call ie.Wait(2000) '�w��~���b��~
                ie.WaitLoading '�ǂݍ��݂��I���܂őҋ@

                Dim temp: temp = ie.GetIEVersion  'IE�̃o�[�W�����𕶎���Ŏ擾
                ie.DisableConfirmFunction 'JS �Ŏ������ꂽconfirm�֐�����ɁB�Ăяo�����Ɋm�F�_�C�A���O��\�������Ȃ�
                Call ie.ExecuteJavaScript("window.resizeTo(10,10);") '�C�ӂ�JavaScript�R�[�h�����s

                ie.Quit 'IE�����
        End If
End Sub
