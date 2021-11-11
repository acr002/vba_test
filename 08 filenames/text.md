## �t�@�C�����̎擾
VBA�Ńt�@�C�������擾������@�͂���������܂����A���������Ƃ��悭�g�����̂���Y�^�Ƃ��ċL�q���܂��B

���L�̊֐��Ƀt�H���_���w�肵�ČĂяo���܂��B
�t���p�X��������Collection��Ԃ��܂��B

```VB
Private Function filenames_sub(ByVal a_path As String) As Collection
  Dim fso      As Object
  Dim r_cc     As Collection
  Dim cc       As Collection
  Dim ii       As Variant
  Dim b_file   As Object
  Dim b_folder As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set cc = New Collection
  For Each b_file In fso.getfolder(a_path).Files
    cc.Add b_file.path
  Next b_file
  For Each b_folder In fso.getfolder(a_path).subfolders
    Set r_cc = filenames_sub(b_folder.path)
    For Each ii In r_cc
      cc.Add ii
    Next ii
  Next b_folder
  Set fso = Nothing
  Set filenames_sub = cc
End Function
```

## ���̊֐��ɂ���
�v���C�x�[�g�Ȋ֐��Ƃ��Ă��܂��B
�ŏ��̓N���X�ɂ��Ă����̂ł����A�ڐA�����ʓ|�Ȃ̂Ŋ֐��ɂ��܂����B

�g�p����@�\��FileSystemObject�ł��B
�Ăяo���͎Q�Ɛݒ�ł͂Ȃ�Createobject�ɂ��Ă��܂��B
���x���z�z���l�����܂����B

�ċA�I�ɏ������܂��B
�t�@�C����(�t���p�X)��Collection�Ɋi�[���A����Collection��Ԃ��܂��B
Collection���g�p�������Ƃ̂Ȃ����͐g�\���Ă��܂���������܂��񂪁A���g�͕�����ł��B
�֐�����̖߂�l���󂯎��ۂ�Set�X�e�[�g�����g���g�p���܂�(���܂��Ȃ��Ǝv���Ă�������)�B

## �Ăяo����
�֐��͌Ăяo���Ďg�p���܂��̂ŁA�Ăяo�������K�v�ł��B
�܂��A���̊֐��̈����͑Ώۂ̃t�H���_�ł��B�u�ǂ̃t�H���_�𒲂ׂ邩�v���w�肷��Ƃ������Ƃł��B
�t�H���_�̎w����@�ɂ��āA�Q�T���v�������܂����B

### �Ώۃt�H���_���R�[�h���Ŏw�肷����@
���̃}�N������̃G�N�Z�����u���Ă���t�H���_�Ɂuin�v�Ƃ����t�H���_������Ɖ��肵�܂��B
(�uin�v�t�H���_���Ȃ��ƃG���[�ɂȂ�܂��B)
�G�N�Z�����u���Ă���Ƃ����`ThisWorkbook.path`�Ŏ擾���܂��B
���̎擾�����p�X��`\in\`��ǉ����Ċ����ł��B
(in�̂Ƃ�������R�ɕς��Ďg�p���Ă�������)

���͂��̕��@���g�p���Ă��܂��B
�Ƃ肠�����uin�v�Ƃ����t�H���_������āA���̒��Ƀt�H���_���Ƃł����̂őΏۂ̃t�@�C�������Ă��܂��΂�������ł��B
�uin�v�t�H���_�Ƃ������[�����݂�����̂ł���Ύg������͂����Ǝv���܂��B

```VB
Public Sub sample1_on_code()
  Dim fns  As Collection
  Dim path As String
  path = ThisWorkbook.path & "\in\"     '�w�肷��t�H���_
  Set fns = filenames_sub(path)     '�֐����Ăяo����fns�Ɍ��ʂ����܂�
End Sub
'-----------------------------------------------------------------------------
```

### �_�C�A���O���g��
�}�N�������s����l���N���킩��Ȃ��ꍇ��A�ǂ�ȏ󋵂Ŏg�p����邩�킩��Ȃ��ꍇ�̂��߂̃_�C�A���O���g�����@�ł��B
�}�N�������s����ƁA�G�N�Z��������t�H���_����ɂ��ă_�C�A���O���J���܂��B
�t�H���_���w�肵�Ă��炦�΁A�֐��Ɏw�肳�ꂽ�t�H���_��n���܂��B
�t�H���_���w�肳��Ȃ��������̃t�H���_���֐��ɓn���܂��B

```VB
Public Sub sample2_dialog()
  Dim path_this As String
  Dim path      As String
  Dim fns       As Collection
  path_this = ThisWorkbook.path & "\"
  With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = path_this
    .InitialView = msoFileDialogViewDetails
    .Title = "�w�t�H���_�x��I��ł�������"
    If .Show = True Then
      path = .SelectedItems(1)
    Else
      path = path_this
    End If
  End With
  Set fns = filenames_sub(path)
End Sub
```


## ���̑��̕��@
�u�b�N���J���ۂɂ̓t�@�C�������K�v�ƂȂ�܂��B
�������Ώۂɂ���Ȃ�VBA�̋L�q���ɒ��ڏ���������̎�ł��B
�擾�ł͂Ȃ��A�w��ł������Ԃ��Ȃ�������A���������g��Ȃ��ꍇ�͂��̕��@�ł������Ǝv���܂��B

�~�����t�@�C����������������ꍇ�́A���C���h�J�[�h���g����Dir�֐����֗��ł��B
���Ɉ�̃t�H���_�������ŁA�K�w���Ȃ��悤�ȏ����̏ꍇ�́A�C�y��������ɂȂ�Ǝv���܂��B
�����ADir�֐��ɂ͒��ӓ_������܂��B����͎��̃t�@�C�������擾����ɂ́u�������ȗ����ČĂяo���v�Ƃ������Ƃł��B
���C�����[�`����Dir�֐����g�p���A��������񂾐�̃��[�`���ł�Dir�֐����g�������Ǝv���������w�肵�Ă��܂��ƁA���C�����[�`����Dir�֐����ω����Ă��܂��܂��B
���܂����������@�����邩������܂��񂪁A���͒m��Ȃ��̂ŁADir�֐��Ƃ͋���������Ă��܂��B
�܂��ADo loop���g��Ȃ���΂Ȃ�Ȃ����Ƃ�Dir�֐����g��Ȃ����R�̈�ł��B

