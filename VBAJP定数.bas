Attribute VB_Name = "VBAJP�萔"
Option Explicit

Enum �I�[����
    �㋴ = xlUp
    ���[ = xlDown
    �E�[ = xlToRight
    ���[ = xlToLeft
End Enum

Enum �Z���I����@
    �\���`������ = xlCellTypeAllFormatConditions
    �����ݒ肠�� = xlCellTypeAllValidation
    ��̃Z�� = xlCellTypeBlanks
    �R�����g���� = xlCellTypeComments
    �萔���� = xlCellTypeConstants
    �������� = xlCellTypeFormulas
    �Ō�̃Z�� = xlCellTypeLastCell
    �����\���`�� = xlCellTypeSameFormatConditions
    �������� = xlCellTypeSameValidation
    ���Z�� = xlCellTypeVisible
End Enum

Enum �Z���I�������l
    �G���[�l = xlErrors
    �_���l = xlLogical
    ���l = xlNumbers
    ���� = xlTextValues
End Enum

Enum ���ߕ�
    ���ߕ��W�� = xlFillDefault
    �A���f�[�^ = xlFillSeries
    �R�s�[ = xlFillCopy
    �����̂� = xlFillFormats
    �����Ȃ� = xlFillValues
    �N�P�� = xlFillYears
    ���P�� = xlFillMonths
    ���P�� = xlFillDays
    �T���P�� = xlFillWeekdays
    ���Z = xlLinearTrend
    ��Z = xlGrowthTrend
End Enum

Enum �V�t�g����
    ������ɃV�t�g = xlShiftUp
    �������ɃV�t�g = xlShiftDown
    �E�����ɃV�t�g = xlShiftToRight
    �������ɃV�t�g = xlShiftToLeft
End Enum

Enum �\��t�����@
    ���ׂē\��t�� = xlPasteAll
    �����̂ݓ\��t�� = xlPasteFormulas
    �l�̂ݓ\��t�� = xlPasteValues
    �����̂ݓ\��t�� = xlPasteFormats
    �R�����g�̂ݓ\��t�� = xlPasteComments
    ���͋K���̂ݓ\��t�� = xlPasteValidation
    �r���������S�ē\��t�� = xlPasteAllExceptBorders
    �� = xlPasteColumnWidths
    �����Ɛ��l�̏����̂ݓ\��t�� = xlPasteFormulasAndNumberFormats
    �l�Ɛ��l�̏����̂ݓ\��t�� = xlPasteValuesAndNumberFormats
    ���ׂẴe�[�}��\��t�� = xlPasteAllUsingSourceTheme
    ���ׂĂ̌��������t��������\��t�� = xlPasteAllMergingConditionalFormats
End Enum

Enum �\���`���p�^�[���萔
    �ʉ� = 1
    �����_�ȉ�1�� = 2
    �����_�ȉ�2�� = 3
    ��4���܂�0���� = 4
    ��8���܂�0���� = 5
    ���� = 6
    ����j���t�� = 7
    �a�� = 8
    �a��j���t�� = 9
    ���� = 10
    ������ = 11
    ����AMPM = 12
End Enum

Enum �Z�����ʒu
    ���W�� = xlGeneral
    �����l�� = xlLeft
    ���������� = xlCenter
    ���E�l�� = xlRight
    ���J��Ԃ� = xlFill
    �����[���� = xlJustify
    ���I��͈͓��Œ��� = xlCenterAcrossSelection
    ���ϓ�����t�� = xlDistributed
End Enum

Enum �Z���c�ʒu
    �c��l�� = xlTop
    �c�������� = xlCenter
    �c���l�� = xlBottom
    �c���[���� = xlJustify
    �c�ϓ�����t�� = xlDistributed
End Enum

Enum �Z���p�x
    �p�x30�x = 30
    �p�x45�x = 45
    �p�x60�x = 60
    �p�x90�x = 90
    �p�x�}�C�i�X30�x = -30
    �p�x�}�C�i�X45�x = -45
    �p�x�}�C�i�X60�x = -60
    �p�x�}�C�i�X90�x = -90
    �p�x�c���� = xlVertical
End Enum

'�������enum�ł��Ȃ�
'���萔���Ƃł���
'Enum �W���t�H���g��
'    MSP�S�V�b�N = "�l�r �o�S�V�b�N"
'    MSP���� = "�l�r �o����"
'    MS�S�V�b�N = "�l�r �S�V�b�N"
'    MS����1 = "�l�r ����"
'    Arial = "Arial"
'    ArialBlack = "Arial Black"
'    ���C���I = "���C���I"
'End Enum
Public Const �t�H���g��MSP�S�V�b�N = "�l�r �o�S�V�b�N"
Public Const �t�H���g��MSP���� = "�l�r �o����"
Public Const �t�H���g��MS�S�V�b�N = "�l�r �S�V�b�N"
Public Const �t�H���g��MS���� = "�l�r ����"
Public Const �t�H���g��Arial = "Arial"
Public Const �t�H���g��ArialBlack = "Arial Black"
Public Const �t�H���g�����C���I = "���C���I"

Enum �A���_�[���C���p�^�[�����
    �����Ȃ� = xlUnderlineStyleNone
    ���� = xlUnderlineStyleSingle
    ��d���� = xlUnderlineStyleDouble
    ������v = xlUnderlineStyleSingleAccounting
    ��d������v = xlUnderlineStyleDoubleAccounting
End Enum

Enum �J���[�C���f�b�N�X�p�^�[��
    �C���f�b�N�X�� = 1
    �C���f�b�N�X�� = 2
    �C���f�b�N�X�� = 3
    �C���f�b�N�X�� = 4
    �C���f�b�N�X�� = 5
    �C���f�b�N�X���F = 6
    �C���f�b�N�X�� = 7
    �C���f�b�N�X���F = 8
    �C���f�b�N�X���F = 9
    �C���f�b�N�X�[�� = 10
    �C���f�b�N�X���F = 11
    �C���f�b�N�X���y�F = 12
    �C���f�b�N�X�[�� = 13
    �C���f�b�N�X��2 = 20
    �C���f�b�N�X�D�F = 15
    �C���f�b�N�X�Z���D�F = 16
    �C���f�b�N�X�� = 17
    �C���f�b�N�X��2 = 18
    �C���f�b�N�X�������F = 19
    �C���f�b�N�X������ = 20
    �C���f�b�N�X�[��2 = 21
    �C���f�b�N�X���F = 22
    �C���f�b�N�X��2 = 23
    �C���f�b�N�X������ = 24
    �C���f�b�N�X�Z����2 = 25
    �C���f�b�N�X������2 = 26
    �C���f�b�N�X���F2 = 27
    �C���f�b�N�X���F2 = 28
    �C���f�b�N�X��3 = 29
    �C���f�b�N�X���F2 = 30
    �C���f�b�N�X�[��2 = 31
    �C���f�b�N�X�Z���� = 32
    �C���f�b�N�X�� = 33
    �C���f�b�N�X�������F = 34
    �C���f�b�N�X�������� = 35
    �C���f�b�N�X�������F2 = 36
    �C���f�b�N�X�������F2 = 37
    �C���f�b�N�X�����s���N = 38
    �C���f�b�N�X������3 = 39
    �C���f�b�N�X�������F = 40
    �C���f�b�N�X���F2 = 41
    �C���f�b�N�X�Z�����F = 42
    �C���f�b�N�X������ = 43
    �C���f�b�N�X�Z�����F = 44
    �C���f�b�N�X�����I�����W = 45
    �C���f�b�N�X�I�����W = 46
    �C���f�b�N�X���F3 = 47
    �C���f�b�N�X�D�F2 = 48
    �C���f�b�N�X�Z�����F = 49
    �C���f�b�N�X��3 = 50
    �C���f�b�N�X�Z���D�F2 = 51
    �C���f�b�N�X�Z���D�F3 = 52
    �C���f�b�N�X�Z���I�����W = 53
    �C���f�b�N�X�Z���s���N = 54
    �C���f�b�N�X�Z����3 = 55
    �C���f�b�N�X�Z���D�F4 = 56
    �F�������I�ɐݒ� = xlColorIndexAutomatic
    �C���f�b�N�X�Ȃ� = xlColorIndexNone
End Enum

Enum �r���ʒu
    �㋴�̌r�� = xlEdgeTop
    ���[�̌r�� = xlEdgeBottom
    ���[�̌r�� = xlEdgeLeft
    �E�[�̌r�� = xlEdgeRight
    �����̉��� = xlInsideHorizontal
    �����̏c�� = xlInsideVertical
    �E������̎΂ߐ� = xlDiagonalDown
    �E�オ��̎΂ߐ� = xlDiagonalUp
End Enum

Enum �r������
    �׎��� = xlContinuous
    �j�� = xlDash
    ��_���� = xlDashDot
    ��_���� = xlDashDotDot
    �_�� = xlDot
    ��d�� = xlDouble
    �΂ߔj�� = xlSlantDashDot
    ���Ȃ� = xlLineStyleNone
End Enum

Enum �r���̑���
    �ɍ� = xlHairline
    �ׂ� = xlThin
    �� = xlMedium
    ���� = xlThick
End Enum

Enum �Z���w�i�F�p�^�[��
    �h��Ԃ� = xlPatternSolid
    �D�F75�p�[�Z���g = xlGray75
    �D�F50�p�[�Z���g = xlGray50
    �D�F25�p�[�Z���g = xlGray25
    �D�F16�p�[�Z���g = xlGray16
    �D�F8�p�[�Z���g = xlGray8
    ���� = xlHorizontal
    �c�� = xlVertical
    �E������΂ߐ� = xlDown
    �E�オ��΂ߐ� = xlUp
    �`�F�b�N = xlChecker
    �D�F�i�q = xlSemiGray75
    ���א� = xlLightHorizontal
    �c�א� = xlLightVertical
    �E������΂ߍא� = xlLightDown
    �E�オ��΂ߍא� = xlLightUp
    �i�q = xlGrid
    �i�q�א� = xlCrissCross
    ���`�O���f�[�V���� = xlPatternLinearGradient
    ���`�O���f�[�V���� = xlPatternRectangularGradient
End Enum

Enum ��΂����΂��A�h���X�w��
    ��΃A�h���X = True
    ���΃A�h���X = False
End Enum


Enum �t�@�C���쐬�t�H���_���
    ���݂̃t�H���_ = 1
    �}�C�h�L�������g = 2
    �t���p�X = 3
    �w��Ȃ� = 4
End Enum

Enum �I��͈̓p�^�[���w��
    �I��͈̓p�^�[���w��Ȃ� = 0
    �I��͈̓p�^�[�������s = 1
    �I��͈̓p�^�[����s = 2
    �I��͈̓p�^�[�������� = 3
    �I��͈̓p�^�[����� = 4
    �I��͈̓p�^�[���s�X�e�b�v = 5
    �I��͈̓p�^�[����X�e�b�v = 6
End Enum

Public Enum �������؂蕶��
    �Ȃ� = 0
    �J���} = 1
    �^�u = 2
    ���s = 3
    Cr = 4
    ���p�� = 5
    ���̑� = 6
End Enum

Public Enum �z��̒l���w�肵�č폜�I�v�V����
    �S�Y���v�f�폜 = 0
    �ŏ��̗v�f�����폜 = 1
End Enum

Public Enum �t�@�C���������ݕ��@
    �㏑�� = 1
    �A�� = 2
End Enum

Public Enum �e�L�X�g��r���@
    �啶������������� = vbBinaryCompare
    �啶������������ʂ��Ȃ� = vbTextCompare
End Enum
