
# [Excel] �فX�Ɠ��͂���ۂ̐ݒ�

�G�N�Z���œ��͍�Ƃ�����ۂ̐ݒ�ł��B
�ݒ肵�Ȃ��Ă����͂͂ł��܂����A������Ǝ�����Ă�����Ίy�ɂȂ�ꍇ������܂��B
�G�N�Z���̃o�[�W������X�̐ݒ�ɂ���ċ@�\���Ȃ��ꍇ���������܂��̂ł��������������B

### �Z���̕ҏW
�Z���ɕ�����ł����ލہA�u���́v�Ɓu�ҏW�v�Ƃ�����̃��[�h������܂��B
�Z���I���̏�ԂŃL�[�{�[�h�������ĕ���������Ɓu���́v�ɂȂ�܂��B
�Z���I���̏�Ԃ�BackSpace�������ƁA�Z���̒��g���N���A���Ă���u���́v�ɂȂ�܂��B
����A�Z���I���̏�Ԃ�`F2`(�t�@���N�V����2)�������Ɓu�ҏW�v�ɂȂ�܂��B
���͂ƕҏW�̈Ⴂ�́A���͂̏ꍇ�͖��L�[�������Ɓu�Z���̈ړ��v�ɂȂ�܂��B
�ҏW�̏ꍇ�͖��L�[�������Ɓu�Z�����ł̃J�[�\���̈ړ��v�ɂȂ�܂��B
(���L�[�����łȂ�`Home`��`End`�{�^�������l�ł�)
`F2`�{�^���͉������ƂɕҏW�Ɠ��͂�؂�ւ��Ă����̂ŁA�Z�����ړ��ƁA�Z�����̃J�[�\���̈ړ��Ŏg�������邱�Ƃ��ł��܂��B

#### ���ӓ_
IME���I���̏�ԂŁA�ϊ����m�肵�Ă��Ȃ��ꍇ�́A���L�[�������Ă���������̈ړ��ƂȂ�܂��B

### �Z���̈ړ�
TAB�ŉE�̃Z���Ɉړ����܂��B
Enter�ł͈ړ����I�Ԃ��Ƃ��ł��܂��B�I�v�V�����Őݒ肵�Ă��������B
�]�k�ɂȂ�܂����A���̓Z���l���m�肵���琔���o�[�œ��e���m�F�������̂ŁAEnter�������Ă��ړ����Ȃ��l�ɂ��Ă��܂��B

### ���͋K�� (Alt, d, l)
���͋K�������낢�날��܂����A���͍�Ƃ̂Ƃ��͗񂲂ƂɁAIME�́u�I���v�܂��́u�I�t(�p�ꃂ�[�h)�v��ݒ肵�Ă����Ɗy�ɂȂ�܂��B
������A���t�@�x�b�g�Ȃǔ��p�������͂��Ȃ��ꍇ�̓I�t�ɂ��܂��B
���{��ȂǑS�p����͂���ꍇ�̓I���ɂ��܂��B
����IME�̃I���E�I�t�̓Z�����ړ����Ă����Ƃ��ɍ�p���܂��B�u���p�^�S�p�L�[�v�������ΐؑւ����܂��B
���͋K���̃V���[�g�J�b�g��`Alt, d, l`�ł��B

### �\���^��\��
���͂ɕK�v�̂Ȃ���͔�\���ɂ��Ă����܂��B
���ꂾ���ō�Ƃ��͂��ǂ邱�Ƃ�����܂��B
��̔�\���̃V���[�g�J�b�g��`Alt, o, c, h`�ł��B

### ���̍s�ւ̈ړ�
VBA�̃C�x���g�@�\���g���Ď��̍s�ֈړ�����悤�ɂ���ƁA�A���������͂��\�ɂȂ�܂��B

```VBA
' �V�[�g���W���[���ɐݒu���Ă��������B
' 17��ڈȍ~���I�����ꂽ�ꍇ�͎��̍s�̐擪�ɃJ�[�\�����ړ����܂��B
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
  If Target.Column >= 17 Then
    Cells(Target.Row + 1, 1).Select
  End If
End Sub
```

���̃}�N���͑Ώۂ̃V�[�g�̃��W���[���ɒu���Ă��������B
�Z���̑I�����ύX���ꂽ�ۂɂ��̃C�x���g�͔������܂��B
If�`End If�̊Ԃɂ��낢��ǉ����邱�Ƃ��ł��܂��B
`thisworkbook.save`��ǉ�����Εۑ������Ă���܂��B
Target.Row�œ��͂��Ă����s�ԍ����擾�ł���̂ŁA���͒l�̊m�F�Ȃǂ��ł��܂��B

### ���Z
#### ���͂ɓ���������
���{��̏ꍇ�͊����̕ϊ��Ɏ��Ԃ�������܂��B
���l�݂̂��������������͂͑����Ȃ�܂��B
�͂��E�������Ȃ�A�͂���1�œ��͂��A��������2�œ��͂��܂��B
�j���E�����Ȃ�j����1�ɁA������2�Ɠ��͂��܂��B���͂���ۂɔ]�ŕϊ������Ă��܂������ł��B
�I����������ł���ꍇ�Ȃ炠�܂�~�X�Ȃ����͂ł���Ǝv���܂��B
(�I�����͂T���炢�܂ł�����ł��B����ȏ�ɂȂ�ƃ~�X�������A���Ԃ�������Ǝv���܂�)
�G�N�Z���ɂ̓I�[�g�R���v���[�g�@�\(������݂̂ł����A���͗��������Ƃɓ��͌���\�����Ă����)������܂��̂�
�ŏ��́u1 �j���v��u2 �����v�ȂǂƓ��͂��Ă�����1�Ɠ��͂����`1 �j��`�����Ɍ���܂��B
�܂��A���͋K���Ń��X�g�ɂ��Ă������Ƃ��l�����܂��B`Alt + ��`�Ń��X�g���\������܂��B




���}�̂̃f�[�^��d�q�f�[�^�ɕς���͖̂ʓ|�ł��B
���������Ȃ���΂Ȃ�Ȃ��󋵂����X����܂��B
�����g���ĊȒP�ȃA���P�[�g�������Ƃ���ƌv�������Ƃ��ȂǁA�W�v����O�ɂ́u���́v���K�v�ɂȂ�܂��B
�G�N�Z�����g���ē��͂��Ă����΁A�W�v�ɂȂ���̂��y�ɂȂ�܂��B
���͍�Ƃ��u���y�ɂȂ�悤�Ɂv���邽�߂̃G�N�Z���̐ݒ�ł��B

### �G�N�Z���̊�{�ݒ�
�G�N�Z���͍ŏ�i�ƍ��[�����܂��Ă���Ε\�ƔF�����Ă���܂��B
���̂��߂Ƀw�b�_�ƍ��[�͋�Z��(���������Ă��Ȃ��Z��)���Ȃ��悤�ɂ��܂��B
�w�b�_�͊�{�I�ɂP�s�ɂ��܂��傤�B�I�[�g�t�B���^����בւ��̂Ƃ��ɕs���v�f���Ȃ��Ȃ�܂��B
���[�ɂ͘A��(SEQ�AID�ANo.�Ȃ�)�ɂ��܂��B






