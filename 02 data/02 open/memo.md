
Excel VBA����G�N�Z���t�@�C��(�u�b�N)���������@�ł��B

### �I�u�W�F�N�g
`Workbook�I�u�W�F�N�g`���u�b�N��\���Ă��܂��B
���̃I�u�W�F�N�g���Ǘ����Ă���̂�`Workbooks�I�u�W�F�N�g`�ł��B
�P��̍Ō��`s`���t���������ł��BWorkbook�̕����`�ł��B
�Ǘ����Ă���Ƃ����܂������A������R���e�i�ł��B
�J���Ă���u�b�N�S�Ă��i�[����Ă��܂��B
���ꂩ��J���u�b�N�ɑ΂��Ă����̃I�u�W�F�N�g����Ăяo���܂��B

### ���ʂȃu�b�N
��L��Workbooks�I�u�W�F�N�g���g�킸�ɃA�N�Z�X�ł���u�b�N������܂��B
�܂��AVBA(���s���Ă���}�N��)�������Ă���u�b�N�ŁA`ThisWorkbook`�ŕ\���܂��B
������A�A�N�e�B�u�ɂȂ��Ă���u�b�N��`ActiveWorkbook`�ŕ\�����Ƃ��ł��܂��B

`ThisWorkbook.Path`�Ńp�X���擾�ł��܂�(��_�ƂȂ�̂ł悭�g���܂�)�B
`ThisWorkbook.Name`�Ńt�@�C�������擾�ł��܂�(�܂��A�g���܂��񂯂�)�B
ActiveWorkbook�͂��܂�g���܂���B���L���Ă��Ȃ��ꍇ��ActiveWorkbook���⊮����邱�Ƃ����R�̈�A����ɁAExcel�ł͐V���ɊJ�����u�b�N��������ActiveWorkbook�ɂȂ��Ă��܂����Ƃ���A���s���ɂ���ĈӐ}���Ȃ����ʂɂȂ�\�������邩��ł��B

### �u�b�N���J��
Workbooks(Workbook�̊Ǘ���)��`.Add`��`.open`���\�b�h���g���ău�b�N���J���܂��B
�����̃��\�b�h��`�J�����u�b�N`��Ԃ��̂ŁA`set + �ϐ� =`�Ŏ󂯎��܂��B

#### �V�K�u�b�N

```vb
Dim wb As Workbook
Set wb = Workbooks.Add
```

#### �����̃u�b�N
�����̃u�b�N���J���ɂ̓p�X���K�v�ł��B

```vb
Dim wb As Workbook
Set wb = Workbooks.Open(Filename:=�p�X)
```

�p�X�͑��΃p�X�ł���΃p�X�ł��g�����Ƃ��ł��܂��B

#### ���O��t���ĕۑ�
`SaveAs`���\�b�h���g���܂��B�����Ƀp�X(�t�@�C����)���w�肵�܂��B

```vb
wb.SaveAs Filename:=�p�X
```

#### ����
`Close`���\�b�h���g���܂��B
�l�I�ɂ�`savechanges:=false`��t���邱�Ƃ������ł�(�قڑS��)�B
���̈����́u�ύX���邯�Ǖۑ����Ȃ��Ă����́H�v�Ƃ����G�N�Z������̃��b�Z�[�W���o���Ȃ��悤�ɂ�����̂ł��B
���O�ɕۑ����Ă���Ζ�肠��܂���B

```vb
wb.Close savechanges:=False
```

### ���Z
#### �J�����x���グ��
�G�N�Z���̓u�b�N���J�����ہA���낢��Ȃ��Ƃ�����Ă���悤�ł��B
�Čv�Z���Ă݂���A��ʕ`�ʂ��v�Z������B
�u�b�N���J�������̓u�b�N�̗e�ʂɂ���Ă��ς��܂��B
�o���邾������������@������������܂��B

- ��ʕ`�ʂ��~�߂�
- �C�x���g���~�߂�
- �Čv�Z���~�߂�

```vb
' �~�߂�
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

' �����ɏ������L�q

' ���ɖ߂�
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.ScreenUpdating = True
```

�ꍇ�ɂ���Ă͍Čv�Z���K�v�ȂƂ�������܂��B
�C�x���g�̂��Ƃ͂����Y��Ă��܂��܂��B
�Ȃ̂Ń}�N���������Ă���̂�����K�v���Ȃ���Ή�ʕ`��(ScreenUpdating)���~�߂邾���ł��Ȃ葬���Ȃ�܂��B

##### ReadOnly��t���ď������������J���A�����Ĉ��S���𓾂�
�ǂݍ��ݐ�p�ŊJ�����ƂŁA�ق�̏��������J����悤�ł��B
�܂��A�f�[�^��ύX���Ă��܂�����������Ȃ��Ƃ����s���������Ă���܂��B

Open���\�b�h�̈�����ReadOnly:=True��ǉ����܂��B

```vb
Set wb = Workbooks.open(Filename:=�p�X, ReadOnly:=True)
```

### �����
����Ƃ��ẮA���s�t�@�C��(ThisWorkbook)�͎蓮�ŊJ���܂��B
�ȍ~�̓}�N���ŁA�f�[�^��ۊǂ���t�@�C�����J��(Add�ō��)�A
�f�[�^�̌��������Ă���t�@�C�����J���āA�ۊǃt�@�C���Ƀf�[�^���ڂ����肵�܂��B
���t�@�C������������ꍇ�́A�����ς݂̃t�@�C���͕��Ă���A�V�����t�@�C�����J���܂��B
�Ȃ̂œ����ɊJ���Ă���̂͂R�t�@�C���Ƃ������ƂɂȂ�܂��B
�����A����t�@�C����A���ŊJ���ā`���ā`�A �J���ā`���ā`�Ƃ���Ă����ƁA���������������Ȃ邱�Ƃ�����܂��B
���̌o���ł�5,000�t�@�C���𒴂���ƃA�N�e�B�x�[�g���ς������A���������Ԃ����������肵�܂��B
�ʓ|�ł����A2,000�t�@�C���ʂɕ������ď������邱�Ƃ��������߂��܂��B


