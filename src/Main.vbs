'==============================================================================
' �A�v���P�[�V�����̃G���g���[�|�C���g
'==============================================================================
Sub Main()

	'b-PAC�I�u�W�F�N�g�𐶐�
	Set ObjDoc = CreateObject("bpac.Document")

	'P-touch Editor�ō쐬�����e���v���[�g�t�@�C�����I�[�v��
	'lbx�t�@�C����VBS�t�@�C���Ɠ����t�H���_�ɒu���Ă���
	bRet = ObjDoc.Open("..\DATA\template.lbx")

	If bRet Then '����ɃI�[�v���H

		'�Z��
		ObjDoc.GetObject("AddressText").Text = "999-9999 �����s���c�撆�� 1-1-1 �e�X�g�r��"

		'���Ȃ܂�
		ObjDoc.GetObject("NameText").Text = "Jimi Hendrix �l"

		'��������s
		ObjDoc.StartPrint "DocumentName", bpoAutoCut
		ObjDoc.PrintOut 1, bpoAutoCut
		ObjDoc.EndPrint

		'��������s(2)
		' ObjDoc.StartPrint "DocumentName", bpoAutoCut
		' ObjDoc.PrintOut 1, bpoAutoCut
		' ObjDoc.EndPrint

		'��������s(3)
		' ObjDoc.StartPrint "DocumentName", bpoAutoCut
		' ObjDoc.PrintOut 1, bpoAutoCut
		' ObjDoc.EndPrint

		ObjDoc.Close

	End If

	'b-PAC�I�u�W�F�N�g�����
	Set ObjDoc = Nothing

End Sub

Call Main()
