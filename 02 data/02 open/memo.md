
Excel VBA����G�N�Z���t�@�C��(�u�b�N)���J�����@�ł��B

### �I�u�W�F�N�g
`workbook�I�u�W�F�N�g` ���u�b�N��\���Ă��܂��B
���̃I�u�W�F�N�g���Ǘ����Ă���̂� `workbooks�I�u�W�F�N�g` �ł��B
�P��̍Ō��`s`���t���������ł��Bworkbook�̕����`�ł��B
�Ǘ����Ă���Ƃ����܂������A������R���e�i�ł��B
�J���Ă���u�b�N�S�Ă��i�[����Ă��܂��B
���ꂩ��J���u�b�N�ɑ΂��Ă����̃I�u�W�F�N�g����Ăяo���܂��B

### ���ʂȃu�b�N
��L��workbooks�I�u�W�F�N�g���g�킸�ɃA�N�Z�X�ł���u�b�N������܂��B
�܂��AVBA(���s���Ă���}�N��)�������Ă���u�b�N�ŁA `Thisworkbook` �ŕ\���܂��B
������A�A�N�e�B�u�ɂȂ��Ă���u�b�N�� `Activeworkbook` �ŕ\�����Ƃ��ł��܂��B

`Thisworkbook.path`�Ńp�X���擾�ł��܂�(��_�ƂȂ�̂ł悭�g���܂�)�B
`Thisworkbook.name`�Ńt�@�C�������擾�ł��܂�(�܂��A�g���܂��񂯂�)�B
Activeworkbook�͂��܂�g���܂���B���L���Ă��Ȃ��ꍇ��Activeworkbook���⊮����邱�Ƃ����R�̈�A����ɁAExcel�ł͐V���ɊJ�����u�b�N��������Activeworkbook�ɂȂ��Ă��܂����Ƃ���A���s���ɂ���ĈӐ}���Ȃ����ʂɂȂ�\�������邩��ł��B

### �u�b�N���J��
����workbook�̊Ǘ���(workbooks)��`.Add`��`.open`���\�b�h���g���ău�b�N���J���܂��B
�����̃��\�b�h��`�J�����u�b�N`��Ԃ��̂ŁA`set + �ϐ� =`�Ŏ󂯎��܂��B

#### �V�K�u�b�N

```vba
dim wb as workbook
set wb = workbooks.add
```

#### �����̃t�@�C��
�����̃t�@�C�����J���ɂ̓p�X���K�v�ł��B

```vba
dim wb as workbook
set wb = workbooks.open(filename:=�p�X)
```

�p�X�͑��΃p�X�ł���΃p�X�ł��g�����Ƃ��ł��܂��B

#### ���O��t���ĕۑ�
`saveas`���\�b�h���g���܂��B�����Ƀp�X(�t�@�C����)���w�肵�܂��B

```vba
wb.saveas filename:=�p�X
```

#### ����
`close`���\�b�h���g���܂��B
�l�I�ɂ�`savechanges:=false`��t���邱�Ƃ������ł�(�قڑS��)�B
���̈����́u�ύX���邯�Ǖۑ����Ȃ��Ă����́H�v�Ƃ����G�N�Z������̃��b�Z�[�W���o���Ȃ��悤�ɂ�����̂ł��B
���O�ɕۑ����Ă���Ζ�肠��܂���B

```vba
wb.close savechanges:=false
```

### �����
�G�N�Z��VBA�͂��낢��ȕ����爫�҈�������Ă��܂��B�܂���������ꂽ������܂��B
�}�N��(VBA)�������Ă���Ɗg���q��xlsm�ɂ���K�v������A�Z�L�����e�B�ɂ���Ă͎��s�ł��Ȃ��Ȃ�����A���[���ł͎�拑�ۂ��ꂽ�肵�܂��B
�Ȃ̂ŁA�Ђ�����Ǝg�����߂ɁA�f�[�^�������Ă���t�@�C���ƁA�}�N���������Ă���t�@�C���𕪂��܂��B
�ʂ̃t�@�C���ɂ���͖̂ʓ|�ł����A�ȊO�ƊǗ������₷�������肷��̂ł������߂ł��B

���āA�f�[�^�p�u�b�N���}�N���ł�����̂ŁA�G�N�Z���ŊJ���K�v������܂��B
(`�u�b�N == �G�N�Z���̃t�@�C��`�ł��B)
�G�N�Z����VBA����u�b�N���J�����@�����L�ɋL�q���܂��B
�܂��̓G�N�Z���̃u�b�N�I�u�W�F�N�g����������܂��B











