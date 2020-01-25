Attribute VB_Name = "VBAEN_CONST"
Option Explicit

Enum OrientForEnding '�I�[����
    vbeEndUp = xlUp
    vbeEndDown = xlDown
    vbeEndRight = xlToRight
    vbeEndLeft = xlToLeft
End Enum

Enum SelectionMethodForCell '�Z���I����@
    vbeExistsTypeAllFormatConditions = xlCellTypeAllFormatConditions '�\���`������
    vbeExistsTypeAllValidation = xlCellTypeAllValidation '�����ݒ肠��
    vbeExistsTypeBlanks = xlCellTypeBlanks  '��̃Z��
    vbeExistsTypeComments = xlCellTypeComments '�R�����g����
    vbeExistsTypeConstants = xlCellTypeConstants '�萔����
    vbeExistsTypeFormulas = xlCellTypeFormulas '��������
    vbeExistsTypeLastCell = xlCellTypeLastCell '�Ō�̃Z��
    vbeExistsTypeSameFormatConditions = xlCellTypeSameFormatConditions '�����\���`��
    vbeExistsTypeSameValidation = xlCellTypeSameValidation '��������
    vbeExistsTypeVisible = xlCellTypeVisible '���Z��
End Enum

Enum ConditionsForSelectingCells '�Z���I�������l
    vbeValuesOfErrors = xlErrors '�G���[�l
    vbeValuesOfLogical = xlLogical '�_���l
    vbeValuesOfNumbers = xlNumbers '���l
    vbeValuesOfText = xlTextValues '����
End Enum

Enum FillMethod '���ߕ�
    vbeFillMethodDefault = xlFillDefault '���ߕ��W��
    vbeFillMethodSeries = xlFillSeries '�A���f�[�^
    vbeFillMethodCopy = xlFillCopy '�R�s�[
    vbeFillMethodFormatsOnly = xlFillFormats '�����̂�
    vbeFillMethodValuesOnly = xlFillValues '�����Ȃ�
    vbeFillMethodYears = xlFillYears '�N�P��
    vbeFillMethodMonths = xlFillMonths '���P��
    vbeFillMethodDays = xlFillDays '���P��
    vbeFillMethodWeekdays = xlFillWeekdays  '�T���P��
    vbeFillMethodLinearTrend = xlLinearTrend '���Z
    vbeFillMethodGrowthTrend = xlGrowthTrend '��Z
End Enum

Enum ShiftOrient '�V�t�g����
    vbeShiftForwordToUp = xlShiftUp '������ɃV�t�g
    vbeShiftForwordToDown = xlShiftDown '�������ɃV�t�g
    vbeShiftForwordToRight = xlShiftToRight '�E�����ɃV�t�g
    vbeShiftForwordToLeft = xlShiftToLeft '�������ɃV�t�g
End Enum

Enum PasteMethod '�\��t�����@
    vbePasteAll = xlPasteAll '���ׂē\��t��
    vbePasteOnlyFomurals = xlPasteFormulas '�����̂ݓ\��t��
    vbePasteOnlyValues = xlPasteValues '�l�̂ݓ\��t��
    vbePasteOnlyFormats = xlPasteFormats '�����̂ݓ\��t��
    vbePasteOnlyComments = xlPasteComments '�R�����g�̂ݓ\��t��
    vbePasteOnlyValidation = xlPasteValidation '���͋K���̂ݓ\��t��
    vbePasteAllExceptBoders = xlPasteAllExceptBorders '�r���������S�ē\��t��
    vbePasteOnlyColumnWidths = xlPasteColumnWidths '��
    vbePasteOnlyFormulasAndNumberFormats = xlPasteFormulasAndNumberFormats  '�����Ɛ��l�̏����̂ݓ\��t��
    vbePasteOnlyValuesAndNumberFormats = xlPasteValuesAndNumberFormats '�l�Ɛ��l�̏����̂ݓ\��t��
    vbePasteAllUsingSourceTheme = xlPasteAllUsingSourceTheme '���ׂẴe�[�}��\��t��
    vbePasteAllMergingConditionalFormats = xlPasteAllMergingConditionalFormats  '���ׂĂ̌��������t��������\��t��
End Enum

Enum VisualFormatPatternForCell '�\���`���p�^�[���萔
    vbeCurency = 1 '�ʉ�
    vbeOneDecimalPlace = 2 '�����_�ȉ�1��
    vbeTwoDecimalPlace = 3 '�����_�ȉ�2��
    vbeZeroPaddingUpTo4digit = 4 '��4���܂�0����
    vbeZeroPaddingUpTo8digit = 5 '��8���܂�0����
    vbeAnnoDomini = 6 '����
    vbeAnnoDominiWithDate = 7 '����j���t��
    vbeJapaneseCalendar = 8 '�a��
    vbeJapaneseCalendarWithDate = 9 '�a��j���t��
    vbeDateAndTime = 10 '����
    vbeDareAndTimeAndMinutes = 11 '������
    vbeDateAndTimeWithAMandPM = 12 '����AMPM
End Enum

Enum HorizentalPositionForCell '�Z�����ʒu
    vbeGeneral = xlGeneral '���W��
    vbeLeft = xlLeft '�����l��
    vbeCenter = xlCenter '����������
    vbeRight = xlRight '���E�l��
    vbeFill = xlFill '���J��Ԃ�
    vbeHorizentalJustify = xlJustify '�����[����
    vbeCenterAcrossSelection = xlCenterAcrossSelection '���I��͈͓��Œ���
    vbeHorizentalDistributed = xlDistributed '���ϓ�����t��
End Enum

Enum VerticalPositionForCell '�Z���c�ʒu
    vbeTop = xlTop '�c��l��
    vbeCenter = xlCenter '�c��������
    vbeBottom = xlBottom '�c���l��
    vbeVerticalJustify = xlJustify '�c���[����
    vbeVerticalDistributed = xlDistributed '�c�ϓ�����t��
End Enum

Enum CellDegree  '�Z���p�x
    vbeDegree30 = 30 '�p30�x
    vbeDegree45 = 45 '�p�x45�x
    vbeDegree60 = 60 '�p�x60�x
    vbeDegree90 = 90 '�p�x90�x
    vbeDegreeMinus30 = -30 '�p�x�}�C�i�X30�x
    vbeDegreeMinus45 = -45 '�p�x�}�C�i�X45�x
    vbeDegreeMinus60 = -60 '�p�x�}�C�i�X60�x
    vbeDegreeMinus90 = -90 '�p�x�}�C�i�X90�x
    vbeDegreeVertical = xlVertical '�p�x�c����
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
Public Const FontNameMSPGotic = "�l�r �o�S�V�b�N" '�t�H���g��MSP�S�V�b�N
Public Const FontNameMSPMincho = "�l�r �o����" '�t�H���g��MSP����
Public Const FontNameMSGotic = "�l�r �S�V�b�N" '�t�H���g��MS�S�V�b�N
Public Const FontNameMSMincho = "�l�r ����" '�t�H���g��MS����
Public Const FontNameArial = "Arial" '�t�H���g��Arial
Public Const FontNameArialBlack = "Arial Black" '�t�H���g��ArialBlack
Public Const FontName���C���I = "���C���I" '�t�H���g�����C���I

Enum StyleOfUnderLinePattern '�A���_�[���C���p�^�[�����
    vbeUnderlineStyleNone = xlUnderlineStyleNone '�����Ȃ�
    vbeUnderlineStyleSingle = xlUnderlineStyleSingle '����
    vbeUnderlineStyleDouble = xlUnderlineStyleDouble '��d����
    vbeUnderlineStyleSingleAccounting = xlUnderlineStyleSingleAccounting '������v
    vbeUnderlineStyleDoubleAccounting = xlUnderlineStyleDoubleAccounting '��d������v
End Enum

'Enum PatternOfColorIndex '�J���[�C���f�b�N�X�p�^�[��
'    IndexBlack = 1 '�C���f�b�N�X��
'    �C���f�b�N�X�� = 2 '�C���f�b�N�X��
'    �C���f�b�N�X�� = 3 '�C���f�b�N�X��
'    �C���f�b�N�X�� = 4 '�C���f�b�N�X��
'    �C���f�b�N�X�� = 5 '�C���f�b�N�X��
'    �C���f�b�N�X���F = 6 '�C���f�b�N�X���F
'    �C���f�b�N�X�� = 7 '�C���f�b�N�X��
'    �C���f�b�N�X���F = 8 '�C���f�b�N�X���F
'    �C���f�b�N�X���F = 9 '�C���f�b�N�X���F
'    �C���f�b�N�X�[�� = 10 '�C���f�b�N�X�[��
'    �C���f�b�N�X���F = 11 '�C���f�b�N�X���F
'    �C���f�b�N�X���y�F = 12 '�C���f�b�N�X���y�F
'    �C���f�b�N�X�[�� = 13 '�C���f�b�N�X�[��
'    �C���f�b�N�X��2 = 20 '�C���f�b�N�X��2
'    �C���f�b�N�X�D�F = 15 '�C���f�b�N�X�D�F
'    �C���f�b�N�X�Z���D�F = 16 '�C���f�b�N�X�Z���D�F
'    �C���f�b�N�X�� = 17 '�C���f�b�N�X��
'    �C���f�b�N�X��2 = 18 '�C���f�b�N�X��2
'    �C���f�b�N�X�������F = 19 '�C���f�b�N�X�������F
'    �C���f�b�N�X������ = 20 '�C���f�b�N�X������
'    �C���f�b�N�X�[��2 = 21 '�C���f�b�N�X�[��2
'    �C���f�b�N�X���F = 22 '�C���f�b�N�X���F
'    �C���f�b�N�X��2 = 23 '�C���f�b�N�X��2
'    �C���f�b�N�X������ = 24 '�C���f�b�N�X������
'    �C���f�b�N�X�Z����2 = 25 '�C���f�b�N�X�Z����2
'    �C���f�b�N�X������2 = 26 '�C���f�b�N�X������2
'    �C���f�b�N�X���F2 = 27 '�C���f�b�N�X���F2
'    �C���f�b�N�X���F2 = 28 '�C���f�b�N�X���F2
'    �C���f�b�N�X��3 = 29 '�C���f�b�N�X��3
'    �C���f�b�N�X���F2 = 30 '�C���f�b�N�X���F2
'    �C���f�b�N�X�[��2 = 31 '�C���f�b�N�X�[��2
'    �C���f�b�N�X�Z���� = 32 '�C���f�b�N�X�Z����
'    �C���f�b�N�X�� = 33 '�C���f�b�N�X��
'    �C���f�b�N�X�������F = 34 '�C���f�b�N�X�������F
'    �C���f�b�N�X�������� = 35 '�C���f�b�N�X��������
'    �C���f�b�N�X�������F2 = 36 '�C���f�b�N�X�������F2
'    �C���f�b�N�X�������F2 = 37 '�C���f�b�N�X�������F2
'    �C���f�b�N�X�����s���N = 38 '�C���f�b�N�X�����s���N
'    �C���f�b�N�X������3 = 39 '�C���f�b�N�X������3
'    �C���f�b�N�X�������F = 40 '�C���f�b�N�X�������F
'    �C���f�b�N�X���F2 = 41 '�C���f�b�N�X���F2
'    �C���f�b�N�X�Z�����F = 42 '�C���f�b�N�X�Z�����F
'    �C���f�b�N�X������ = 43 '�C���f�b�N�X������
'    �C���f�b�N�X�Z�����F = 44 '�C���f�b�N�X�Z�����F
'    �C���f�b�N�X�����I�����W = 45 '�C���f�b�N�X�����I�����W
'    �C���f�b�N�X�I�����W = 46 '�C���f�b�N�X�I�����W
'    �C���f�b�N�X���F3 = 47 '�C���f�b�N�X���F3
'    �C���f�b�N�X�D�F2 = 48 '�C���f�b�N�X�D�F2
'    �C���f�b�N�X�Z�����F = 49 '�C���f�b�N�X�Z�����F
'    �C���f�b�N�X��3 = 50 '�C���f�b�N�X��3
'    �C���f�b�N�X�Z���D�F2 = 51 '�C���f�b�N�X�Z���D�F2
'    �C���f�b�N�X�Z���D�F3 = 52 '�C���f�b�N�X�Z���D�F3
'    �C���f�b�N�X�Z���I�����W = 53 '�C���f�b�N�X�Z���I�����W
'    �C���f�b�N�X�Z���s���N = 54 '�C���f�b�N�X�Z���s���N
'    �C���f�b�N�X�Z����3 = 55 '�C���f�b�N�X�Z����3
'    �C���f�b�N�X�Z���D�F4 = 56 '�C���f�b�N�X�Z���D�F4
'    �F�������I�ɐݒ� = xlColorIndexAutomatic '�F�������I�ɐݒ�
'    �C���f�b�N�X�Ȃ� = xlColorIndexNone '�C���f�b�N�X�Ȃ�
'End Enum

Enum PatternOfColorIndex '�J���[�C���f�b�N�X�p�^�[��
   vbeIndexBlack = 1 '�C���f�b�N�X��
   vbeIndexWhite = 2 '�C���f�b�N�X��
   vbeIndexRed = 3 '�C���f�b�N�X��
   vbeIndexGreen = 4 '�C���f�b�N�X��
   vbeIndexBlue = 5 '�C���f�b�N�X��
   vbeIndexYellow = 6 '�C���f�b�N�X���F
   vbeIndexPurple = 7 '�C���f�b�N�X��
   vbeIndexLightBlue = 8 '�C���f�b�N�X���F
   vbeIndexBrown = 9 '�C���f�b�N�X���F
   vbeIndexDarkGreen = 10 '�C���f�b�N�X�[��
   vbeIndexIndigoBlue = 11 '�C���f�b�N�X���F
   vbeIndexOcher = 12 '�C���f�b�N�X���y�F
   vbeIndexDeepPurple = 13 '�C���f�b�N�X�[��
   vbeIndexGreen2 = 20 '�C���f�b�N�X��2
   vbeIndexGray = 15 '�C���f�b�N�X�D�F
   vbeIndexDarkGray = 16 '�C���f�b�N�X�Z���D�F
   vbeIndexBlueViolet = 17 '�C���f�b�N�X��
   vbeIndexPurple2 = 18 '�C���f�b�N�X��2
   vbeIndexLightYellow = 19 '�C���f�b�N�X�������F
   vbeIndexLightBlue2 = 20 '�C���f�b�N�X������
   vbeIndexDeepPurple2 = 21 '�C���f�b�N�X�[��2
   vbeIndexPeach = 22 '�C���f�b�N�X���F
   vbeIndexBlue2 = 23 '�C���f�b�N�X��2
   vbeIndexLightPurple = 24 '�C���f�b�N�X������
   vbeIndexDarkBlue2 = 25 '�C���f�b�N�X�Z����2
   vbeIndexLightPurple2 = 26 '�C���f�b�N�X������2
   vbeIndexYellow2 = 27 '�C���f�b�N�X���F2
   vbeIndexLightBlue3 = 28 '�C���f�b�N�X���F2
   vbeIndexPurple3 = 29 '�C���f�b�N�X��3
   vbeIndexBrown2 = 30 '�C���f�b�N�X���F2
   vbeIndexDeepGreen2 = 31 '�C���f�b�N�X�[��2
   vbeIndexDeepBlue = 32 '�C���f�b�N�X�Z����
   vbeIndexBlueGreen = 33 '�C���f�b�N�X��
   vbeIndexLightBlue4 = 34 '�C���f�b�N�X�������F
   vbeIndexLightYellowGreen = 35 '�C���f�b�N�X��������
   vbeIndexLightYellow2 = 36 '�C���f�b�N�X�������F2
   vbeIndexLightBlue5 = 37 '�C���f�b�N�X�������F2
   vbeIndexLightPink = 38 '�C���f�b�N�X�����s���N
   vbeIndexLightPurple3 = 39 '�C���f�b�N�X������3
   vbeIndexLightPeach = 40 '�C���f�b�N�X�������F
   vbeIndexIndigoBlue2 = 41 '�C���f�b�N�X���F2
   vbeIndexDeepBlue2 = 42 '�C���f�b�N�X�Z�����F
   vbeIndexLightGreen = 43 '�C���f�b�N�X������
   vbeIndexDarkYellow = 44 '�C���f�b�N�X�Z�����F
   vbeIndexLightOrange = 45 '�C���f�b�N�X�����I�����W
   vbeIndexOrange = 46 '�C���f�b�N�X�I�����W
   vbeIndexIndigoBlue3 = 47 '�C���f�b�N�X���F3
   vbeIndexGray2 = 48 '�C���f�b�N�X�D�F2
   vbeIndexDarkIndigoBlue = 49 '�C���f�b�N�X�Z�����F
   vbeIndexGreen3 = 50 '�C���f�b�N�X��3
   vbeIndexDarkGray2 = 51 '�C���f�b�N�X�Z���D�F2
   vbeIndexDarkGray3 = 52 '�C���f�b�N�X�Z���D�F3
   vbeIndexDarkOrange = 53 '�C���f�b�N�X�Z���I�����W
   vbeIndexDarkPink = 54 '�C���f�b�N�X�Z���s���N
   vbeIndexDarkBlue3 = 55 '�C���f�b�N�X�Z����3
   vbeIndexDarkGray4 = 56 '�C���f�b�N�X�Z���D�F4
   vbeColorIndexAutomatic = xlColorIndexAutomatic '�F�������I�ɐݒ�
   vbeColorIndexNone = xlColorIndexNone '�C���f�b�N�X�Ȃ�
End Enum

Enum BorderPosition '�r���ʒu
    vbeEdgeTop = xlEdgeTop '�㋴�̌r��
    vbeEdgeBottom = xlEdgeBottom '���[�̌r��
    vbeEdgeLeft = xlEdgeLeft '���[�̌r��
    vbeEdgeRight = xlEdgeRight '�E�[�̌r��
    vbeInsideHorizontal = xlInsideHorizontal '�����̉���
    vbeInsideVertical = xlInsideVertical '�����̏c��
    vbeDiagonalDown = xlDiagonalDown '�E������̎΂ߐ�
    vbeDiagonalUp = xlDiagonalUp '�E�オ��̎΂ߐ�
End Enum

Enum StyleOfBoderLine '�r������
    vbeContinuous = xlContinuous '�׎���
    vbeDash = xlDash '�j��
    vbeDashDot = xlDashDot '��_����
    vbeDashDotDot = xlDashDotDot '��_����
    vbeDot = xlDot '�_��
    vbeDouble = xlDouble '��d��
    vbeSlantDashDot = xlSlantDashDot '�΂ߔj��
    vbeLineStyleNone = xlLineStyleNone '���Ȃ�
End Enum

Enum BorderThickness '�r���̑���
    vbeHairline = xlHairline '�ɍ�
    vbeThin = xlThin '�ׂ�
    vbeMedium = xlMedium '��
    vbeThick = xlThick '����
End Enum

Enum CellBackgroundPattern '�Z���w�i�F�p�^�[��
    vbePatternSolid = xlPatternSolid '�h��Ԃ�
    vbeGray75 = xlGray75 '�D�F75�p�[�Z���g
    vbeGray50 = xlGray50 '�D�F50�p�[�Z���g
    vbeGray25 = xlGray25 '�D�F25�p�[�Z���g
    vbeGray16 = xlGray16 '�D�F16�p�[�Z���g
    vbeGray8 = xlGray8 '�D�F8�p�[�Z���g
    vbeHorizontalLine = xlHorizontal '����
    vbeVerticalLine = xlVertical '�c��
    vbeDownBackSlash = xlDown '�E������΂ߐ�
    vbeUpSlash = xlUp '�E�オ��΂ߐ�
    vbeChecker = xlChecker '�`�F�b�N
    vbeSemiGray75 = xlSemiGray75 '�D�F�i�q
    vbeLightHorizontalLine = xlLightHorizontal '���א�
    vbeLightVerticalLine = xlLightVertical '�c�א�
    vbeLightDownBackSlash = xlLightDown '�E������΂ߍא�
    vbeLightUpSlash = xlLightUp '�E�オ��΂ߍא�
    vbeGrid = xlGrid '�i�q
    vbeCrissCross = xlCrissCross '�i�q�א�
    vbePatternLinearGradient = xlPatternLinearGradient '���`�O���f�[�V����
    vbePatternRectangularGradient = xlPatternRectangularGradient '���`�O���f�[�V����
End Enum

Enum AddressDesignation '��΂����΂��A�h���X�w��
    vbeAbsoluteAddress = True
    vbeRelativeAddress = False
End Enum


Enum FolderType '�t�@�C���쐬�t�H���_���
    vbeCurrentFolder = 1 '���݂̃t�H���_
    vbeMyDocument = 2 '�}�C�h�L�������g
    vbeFullPath = 3 '�t���p�X
    vbeNone = 4 '�w��Ȃ�
End Enum

Enum SelectionPattern '�I��͈̓p�^�[���w��
    vbeNonePattern = 0 '�I��͈̓p�^�[���w��Ȃ�
    vbeEvenRows = 1 '�I��͈̓p�^�[�������s
    vbeOddRows = 2 '�I��͈̓p�^�[����s
    vbeEvenCols = 3 '�I��͈̓p�^�[��������
    vbeOddCols = 4 '�I��͈̓p�^�[�����
    vbeRowsByStep = 5 '�I��͈̓p�^�[���s�X�e�b�v
    vbeColsByStep = 6 '�I��͈̓p�^�[����X�e�b�v
End Enum

Public Enum SeparatorChar '�������؂蕶��
    vbeNoneChar = 0 '�Ȃ�
    vbeComma = 1 '�J���}
    vbeTab = 2 '�^�u
    vbeReturn = 3 '���s
    vbeCr = 4 'Cr
    vbeSpaceChar = 5 '���p��
    vbeElseChar = 6 '���̑�
End Enum

Public Enum DeleteByValueOptionForArrayElement '�z��̒l���w�肵�č폜�I�v�V����
    AllMatchValues = 0 '�S�Y���v�f�폜
    FirstMatchValueOnly = 1 '�ŏ��̗v�f�����폜
End Enum

Public Enum FileWriteMethod '�t�@�C���������ݕ��@
    OverWrite = 1 '�㏑��
    SerealNo = 2 '�A��
End Enum

Public Enum TextCompareMode '�e�L�X�g��r���@
    CaseSensitive = vbBinaryCompare
    NoneCaseSensitive = vbTextCompare
End Enum

