Attribute VB_Name = "�֐�"
Public Function �J�X�^��1(���v�͈� As Variant) As Variant
    �J�X�^��1 = �؂�̂�2(���v(���v�͈�), 2)
End Function
Public Function �J�X�^��2(�ΏۃZ�� As Variant) As Variant
    �J�X�^��2 = ��(�ΏۃZ��, 12)
End Function
'���w�̃A�[�N�R�T�C���iarccos�j��x�ŕԂ��֐��ł��B
Public Function �A�[�N�R�T�C���x(cos�l As Variant) As Variant
    �A�[�N�R�T�C���x = Application.WorksheetFunction.Degrees(Application.WorksheetFunction.Acos(cos�l))
End Function

'���w�̃A�[�N�T�C���iarcsin�j��x�ŕԂ��֐��ł��B
Public Function �A�[�N�T�C���x(sin�l As Variant) As Variant
    �A�[�N�T�C���x = Application.WorksheetFunction.Degrees(Application.WorksheetFunction.Asin(sin�l))
End Function

'���w�̃A�[�N�^���W�F���g�iarctan�j���A�x�ŕԂ��֐��ł��B
Public Function �A�[�N�^���W�F���g�x(tan�l As Variant) As Variant
    �A�[�N�^���W�F���g�x = Application.WorksheetFunction.Degrees(Atn(tan�l))
End Function

'���w�̃R�T�C���icos�j��x����������֐��ł��B
Public Function �R�T�C���x(�x As Variant) As Variant
    �R�T�C���x = Cos(Application.WorksheetFunction.Radians(�x))
End Function

'���w�̃T�C���isin�j��x����������֐��ł��B
Public Function �T�C���x(�x As Variant) As Variant
    �T�C���x = Sin(Application.WorksheetFunction.Radians(�x))
End Function

'���w�̃^���W�F���g�x�itan�j��x����������֐��ł��B
Public Function �^���W�F���g�x(�x As Variant) As Variant
    �^���W�F���g�x = Tan(Application.WorksheetFunction.Radians(�x))
End Function

'2�i����10�i���ɕϊ����܂��B
Public Function ��i������\�i��(��i�� As Variant) As Variant

    Dim �\�i���v�Z�p As Variant
    Dim j As Long
    Dim x As Long
    
    �\�i���v�Z�p = 0

    For j = 1 To Len(��i��)
        If Mid(��i��, Len(��i��) - j + 1, 1) = "1" Then
            x = 2 ^ (j - 1)
            �\�i���v�Z�p = �\�i���v�Z�p + x
        End If
    Next j

    ��i������\�i�� = �\�i���v�Z�p

End Function

' n �� m �̎��̗]������߂܂��B
Public Function �]��(�����鐔n As Variant, ���鐔m As Variant) As Variant
    �]�� = �����鐔n Mod ���鐔m
End Function

'16�i����10�i���ɕϊ����܂�
Public Function �\�Z�i������\�i��(�\�Z�i�� As Variant) As Variant
    �\�Z�i������\�i�� = Val("&H" & �\�Z�i��)
End Function

'10�i����2�i���ɕϊ����܂�
Public Function �\�i�������i��(�\�i�� As Variant, Optional ���� As Long = 8) As String
    Dim �r�b�g�t���O As Long
    Dim ��i���v�Z�p As String

    Do Until (�\�i�� < 2 ^ �r�b�g�t���O)
        If (�\�i�� And 2 ^ �r�b�g�t���O) <> 0 Then
            ��i���v�Z�p = "1" & ��i���v�Z�p
        Else
            ��i���v�Z�p = "0" & ��i���v�Z�p
        End If

        �r�b�g�t���O = �r�b�g�t���O + 1
    Loop
    
    Dim n As Long
    Dim padding As String
    For n = 1 To ����
        padding = padding + "0"
    Next n

    �\�i�������i�� = Format(��i���v�Z�p, padding)
End Function

'10�i����16�i���ɕϊ����܂�
Public Function �\�i������\�Z�i��(�\�i�� As Variant, Optional ���� As Long = 4) As Variant
    �\�i������\�Z�i�� = �\�Z�i���p�f�B���O(Hex(�\�i��), "0", ����)
End Function
'�@�\�F�w�蕶�����ߊ֐�
'�����Fstr�@�F�ϊ��O�̕�����
'�@�@�@chr  �F���߂镶��(�P�����ڂ̂ݎg�p)
'�@�@�@digit�F����
'�ߒl�F�w�蕶�����ߌ�̕�����
Private Function �\�Z�i���p�f�B���O(ByVal str As String, _
                     ByVal char As String, _
                     ByVal digit As Long) As String
  Dim tmp As String
  tmp = str
  If Len(str) < digit And Len(char) > 0 Then
    tmp = Right(String(digit, char) & str, digit)
  End If
  �\�Z�i���p�f�B���O = tmp
End Function

'���K�\���̒u���p�^�[����������w�肵�āA���K�\���u�����܂��B
Public Function ���K�\���u��(�����Ώ� As Variant, �u���p�^�[�������� As Variant, �u����̕����� As Variant, Optional �啶������������ As Boolean = False, Optional �ŏ��̈�v���̂ݒu�� As Boolean = False)
    r_RegExp.Pattern = �u���p�^�[��������
    r_RegExp.IgnoreCase = �啶������������
    r_RegExp.Global = Not �ŏ��̈�v���̂ݒu��
    If (IsObject(�����Ώ�)) Then
        ���K�\���u�� = RegEx.Replace(�����Ώ�.Value2, �u����̕�����)
    Else
        ���K�\���u�� = RegEx.Replace(�����Ώ�, �u����̕�����)
    End If
End Function


Public Function �j��(���t�Z�� As Variant, ���1����3 As Variant) As Variant
    �j�� = Application.WorksheetFunction.Weekday(���t�Z��, ���1����3)
End Function
Public Function ������(���l�Z�� As Variant) As Variant
    ������ = Sqr(���l�Z��)
End Function
Public Function ����(���ϔ͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    ���� = Application.WorksheetFunction.Average(���ϔ͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function ������(������ As Variant) As Variant
    ������ = Len(������)
End Function
Public Function �����u��(�u���ΏۃZ�� As Variant, �u���Ώە����� As Variant, �u���㕶���� As Variant) As Variant
    �����u�� = Application.WorksheetFunction.Substitute(�u���ΏۃZ��, �u���Ώە�����, �u���㕶����)
End Function
Public Function ���U(�Z���͈� As Variant) As Variant
    ���U = Application.WorksheetFunction.VarP(�Z���͈�)
End Function
Public Function ��(�����Z�� As Variant) As Variant
    �� = Minute(�����Z��)
End Function
Public Function �s�Ε��U(�Z���͈� As Variant) As Variant
    �s�Ε��U = Application.WorksheetFunction.Var(�Z���͈�)
End Function
Public Function �s�ΕW���΍�(�Z���͈� As Variant) As Variant
    �s�ΕW���΍� = Application.WorksheetFunction.StDev(�Z���͈�)
End Function
Public Function �b(�����Z�� As Variant) As Variant
    �b = Second(�����Z��)
End Function
Public Function �W���΍�(�Z���͈� As Variant) As Variant
    �W���΍� = Application.WorksheetFunction.StDevP(�Z���͈�)
End Function
Public Function �{���؂�グ(���l As Variant, �{��� As Variant) As Variant
    �{���؂�グ = Application.WorksheetFunction.Ceiling(���l, �{���)
End Function
Public Function �{���؂�̂�(���l As Variant, �{��� As Variant) As Variant
    �{���؂�̂� = Application.WorksheetFunction.Floor(���l, �{���)
End Function
Public Function �N(���t�Z�� As Variant) As Variant
    �N = Year(���t�Z��)
End Function
Public Function ���t�ϊ�(�N As Variant, �� As Variant, �� As Variant) As Variant
    ���t�ϊ� = DateSerial(�N, ��, ��)
End Function
Public Function ���t�̍�(��r�P�� As Variant, ���t�Z��1 As Variant, ���t�Z��2 As Variant) As Variant
    ���t�̍� = DateDiff(��r�P��, ���t�Z��1, ���t�Z��2)
End Function
Public Function ��(���t�Z�� As Variant) As Variant
    �� = Day(���t�Z��)
End Function
Public Function �����l(�Z���͈� As Variant) As Variant
    �����l = Application.WorksheetFunction.Median(�Z���͈�)
End Function
Public Function �傫�������牽�Ԗڂ��̒l(�Z���͈� As Variant, ���� As Variant) As Variant
    �傫�������牽�Ԗڂ��̒l = Application.WorksheetFunction.Large(�Z���͈�, ����)
End Function
Public Function �ΐ�(���̐��l As Variant) As Variant
    �ΐ� = Log(���̐��l)
End Function
Public Function �S�p�����𔼊p��(�Ώە����Z�� As Variant) As Variant
    �S�p�����𔼊p�� = Application.WorksheetFunction.Asc(�Ώە����Z��)
End Function
Public Function ��Βl(���l�Z�� As Variant) As Variant
    ��Βl = Abs(���l�Z��)
End Function
Public Function �؂�グ(���l As Variant, �؂�グ�錅�� As Variant) As Variant
    �؂�グ = Application.WorksheetFunction.RoundUp(���l, �؂�グ�錅��)
End Function
Public Function �؂�̂�2(���l As Variant, �؂�̂Ă錅��� As Variant) As Variant
    �؂�̂�2 = Application.WorksheetFunction.RoundDown(���l, �؂�̂Ă錅���)
End Function
Public Function �؂�̂�(���l As Variant) As Variant
    �؂�̂� = Int(���l)
End Function
Public Function ���l�ԃ����_��(�J�n�l As Variant, �I���l As Variant) As Variant
    ���l�ԃ����_�� = Application.WorksheetFunction.RandBetween(�J�n�l, �I���l)
End Function
Public Function ���������[�}������(�ΏۃZ�� As Variant) As Variant
    ���������[�}������ = Application.WorksheetFunction.Roman(�ΏۃZ��)
End Function
Public Function ��������̌���(�J�n�� As Variant, �� As Variant) As Variant
    ��������̌��� = Application.WorksheetFunction.EoMonth(�J�n��, ��)
End Function
Public Function ��������(�J�n�� As Variant, �� As Variant) As Variant
    �������� = Application.WorksheetFunction.EDate(�J�n��, ��)
End Function
Public Function ��p�ΐ�(���̐��l As Variant) As Variant
    ��p�ΐ� = Application.WorksheetFunction.Log10(���̐��l)
End Function
Public Function �����������牽�Ԗڂ��̒l(�Z���͈� As Variant, ���� As Variant) As Variant
    �����������牽�Ԗڂ��̒l = Application.WorksheetFunction.Small(�Z���͈�, ����)
End Function
Public Function ��(�Ώې��l�Z�� As Variant, ���鐔 As Variant) As Variant
    �� = Application.WorksheetFunction.Quotient(�Ώې��l�Z��, ���鐔)
End Function
Public Function ����(���ʒ����Z�� As Variant, �Z���͈� As Variant, �����t���O As Variant) As Variant
    ���� = Application.WorksheetFunction.Rank(���ʒ����Z��, �Z���͈�, �����t���O)
End Function
Public Function �c�\��(�����l As Variant, �����͈� As Variant, ��ԍ� As Variant, Optional �I�v�V����1 As Variant) As Variant
    �c�\�� = Application.WorksheetFunction.VLookup(�����l, �����͈�, ��ԍ�, �I�v�V����1)
End Function
Public Function ���R�ΐ��̒�e�ׂ̂���(�ׂ��ƂȂ鐔 As Variant) As Variant
    ���R�ΐ��̒�e�ׂ̂��� = Exp(�ׂ��ƂȂ鐔)
End Function
Public Function ���R�ΐ�(���̐��l As Variant) As Variant
    ���R�ΐ� = Application.WorksheetFunction.Ln(���̐��l)
End Function
Public Function ���ԕϊ�(�� As Variant, �� As Variant, �b As Variant) As Variant
    ���ԕϊ� = TimeSerial(��, ��, �b)
End Function
Public Function ��(�����Z�� As Variant) As Variant
    �� = Hour(�����Z��)
End Function
Public Function �l�̌ܓ�(���l As Variant, �l�̌ܓ����錅�� As Variant) As Variant
    �l�̌ܓ� = Application.WorksheetFunction.Round(���l, �l�̌ܓ����錅��)
End Function
Public Function �ŕp�l(�Z���͈� As Variant) As Variant
    �ŕp�l = Application.WorksheetFunction.Mode(�Z���͈�)
End Function
Public Function �ő����(���l�͈� As Variant) As Variant
    �ő���� = Application.WorksheetFunction.Gcd(���l�͈�)
End Function
Public Function �ő�(�����͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    �ő� = Application.WorksheetFunction.max(�����͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function �ŏ����{��(���l�͈� As Variant) As Variant
    �ŏ����{�� = Application.WorksheetFunction.Lcm(���l�͈�)
End Function
Public Function �ŏ�(�����͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    �ŏ� = Application.WorksheetFunction.Min(�����͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function ��������(������ As Variant, ������̕����� As Variant) As Variant
    �������� = Left(������, ������̕�����)
End Function
Public Function ���E�󔒕����폜(�Ώە����Z�� As Variant) As Variant
    ���E�󔒕����폜 = Application.WorksheetFunction.Trim(�Ώە����Z��)
End Function
Public Function ��() As Variant
    �� = Now
End Function
Public Function ���v(���v�͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    ���v = Application.WorksheetFunction.Sum(���v�͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function ����(�J�E���g�͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    ���� = Application.WorksheetFunction.Count(�J�E���g�͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function ���X�ϗ����~�����z(���� As Variant, �ϗ����� As Variant, �������z As Variant, �ړI�̐ϗ��z As Variant) As Variant
    ���X�ϗ����~�����z = Application.WorksheetFunction.Pmt(����, �ϗ�����, �������z, �ړI�̐ϗ��z)
End Function
Public Function ���X���[���ԍϊz���̌����ԍϊz(���� As Variant, ���߂���͉̂����ڂ� As Variant, �ԍό��� As Variant, �ؓ����z As Variant, �Ō�Ɏc����z As Variant) As Variant
    ���X���[���ԍϊz���̌����ԍϊz = Application.WorksheetFunction.PPmt(����, ���߂���͉̂����ڂ�, �ԍό���, �ؓ����z, �Ō�Ɏc����z)
End Function
Public Function ���X���[���ԍϊz���̋������z(���� As Variant, ���߂���͉̂����ڂ� As Variant, �ԍό��� As Variant, �ؓ��z As Variant, �Ō�Ɏc����z As Variant) As Variant
    ���X���[���ԍϊz���̋������z = Application.WorksheetFunction.IPmt(����, ���߂���͉̂����ڂ�, �ԍό���, �ؓ��z, �Ō�Ɏc����z)
End Function
Public Function ���X���[���ԍϊz(���� As Variant, �ԍό��� As Variant, �ؓ��z As Variant, �Ō�Ɏc����z As Variant) As Variant
    ���X���[���ԍϊz = Application.WorksheetFunction.Pmt(����, �ԍό���, �ؓ��z, �Ō�Ɏc����z)
End Function
Public Function ��(���t�Z�� As Variant) As Variant
    �� = Month(���t�Z��)
End Function
Public Function �J��Ԃ��\��(�Ώە����� As Variant, �J��Ԃ��� As Variant) As Variant
    �J��Ԃ��\�� = Application.WorksheetFunction.Rept(�Ώە�����, �J��Ԃ���)
End Function
Public Function �ԕ�����(������ As Variant, �擪�����ԍ� As Variant, �����o�������� As Variant) As Variant
    �ԕ����� = Mid(������, �擪�����ԍ�, �����o��������)
End Function
Public Function �p�x(���W�A���Z�� As Variant) As Variant
    �p�x = Application.WorksheetFunction.Degrees(���W�A���Z��)
End Function
Public Function �K��(���̐��l As Variant) As Variant
    �K�� = Application.WorksheetFunction.Fact(���̐��l)
End Function
Public Function �����ڂ̓��t(���t As Variant, �t���O1�܂���2 As Variant) As Variant
    �����ڂ̓��t = Application.WorksheetFunction.WeekNum(���t, �t���O1�܂���2)
End Function
Public Function ���\��(�����l As Variant, �����͈� As Variant, �s�ԍ� As Variant, Optional �I�v�V����1 As Variant) As Variant
    ���\�� = Application.WorksheetFunction.HLookup(�����l, �����͈�, �s�ԍ�, �I�v�V����1)
End Function
Public Function �~����() As Variant
    �~���� = Application.WorksheetFunction.Pi
End Function
Public Function �p�P��̐擪������啶����(�p�P����܂ރZ�� As Variant) As Variant
    �p�P��̐擪������啶���� = Application.WorksheetFunction.Proper(�p�P����܂ރZ��)
End Function
Public Function �p���啶����(�Ώە����Z�� As Variant) As Variant
    �p���啶���� = UCase(�Ώە����Z��)
End Function
Public Function �p����������(�Ώە����Z�� As Variant) As Variant
    �p���������� = LCase(�Ώە����Z��)
End Function
Public Function �c�Ɠ�����(�J�n�� As Variant, �I���� As Variant, �Փ����������Z���͈� As Variant) As Variant
    �c�Ɠ����� = Application.WorksheetFunction.NetworkDays(�J�n��, �I����, �Փ����������Z���͈�)
End Function
Public Function �c�Ɠ�(�J�n�� As Variant, ���� As Variant, �Փ��̓��t���������Z���͈� As Variant) As Variant
    �c�Ɠ� = Application.WorksheetFunction.WorkDay(�J�n��, ����, �Փ��̓��t���������Z���͈�)
End Function
Public Function �E������(������ As Variant, �E����̕����� As Variant) As Variant
    �E������ = Right(������, �E����̕�����)
End Function
Public Function ��v(�����l As Variant, �����͈� As Variant, �ƍ��̎�� As Variant) As Variant
    ��v = Application.WorksheetFunction.Match(�����l, �����͈�, �ƍ��̎��)
End Function
Public Function ���[���ԍϊz�̗��q�������̗݌v�z(���� As Variant, ���[���_�񌎐� As Variant, �ؓ��z As Variant, ��n�� As Variant, ��m�� As Variant, ���������Ȃ�0���񕥂��Ȃ�1 As Variant) As Variant
    ���[���ԍϊz�̗��q�������̗݌v�z = Application.WorksheetFunction.CumIPmt(����, ���[���_�񌎐�, �ؓ��z, ��n��, ��m��, ���������Ȃ�0���񕥂��Ȃ�1)
End Function
Public Function ���[���ԍϊz�̌����������̗݌v�z(���� As Variant, ���[���_�񌎐� As Variant, �ؓ��z As Variant, ��n�� As Variant, ��m�� As Variant, ���������Ȃ�0���񕥂��Ȃ�1 As Variant) As Variant
    ���[���ԍϊz�̌����������̗݌v�z = Application.WorksheetFunction.CumPrinc(����, ���[���_�񌎐�, �ؓ��z, ��n��, ��m��, ���������Ȃ�0���񕥂��Ȃ�1)
End Function
Public Function ��[��ւ񂳂������̂肵�Ԃ�̂邯��������(���� As Variant, ���[���_�񌎐� As Variant, �ؓ��z As Variant, ��n�� As Variant, ��m�� As Variant, ���������Ȃ�0���񕥂��Ȃ�1 As Variant) As Variant
    ��[��ւ񂳂������̂肵�Ԃ�̂邯�������� = Application.WorksheetFunction.CumIPmt(����, ���[���_�񌎐�, �ؓ��z, ��n��, ��m��, ���������Ȃ�0���񕥂��Ȃ�1)
End Function
Public Function ��[��ւ񂳂������̂��񂫂�Ԃ�̂邢��������(���� As Variant, ���[���_�񌎐� As Variant, �ؓ��z As Variant, ��n�� As Variant, ��m�� As Variant, ���������Ȃ�0���񕥂��Ȃ�1 As Variant) As Variant
    ��[��ւ񂳂������̂��񂫂�Ԃ�̂邢�������� = Application.WorksheetFunction.CumPrinc(����, ���[���_�񌎐�, �ؓ��z, ��n��, ��m��, ���������Ȃ�0���񕥂��Ȃ�1)
End Function
Public Function �����_��() As Variant
    �����_�� = Rnd
End Function
Public Function ���W�A��(�p�x�Z�� As Variant) As Variant
    ���W�A�� = Application.WorksheetFunction.Radians(�p�x�Z��)
End Function
Public Function �悱�Ђ傤�т�(�����l As Variant, �����͈� As Variant, �s�ԍ� As Variant, �����̌^ As Variant) As Variant
    �悱�Ђ傤�т� = Application.WorksheetFunction.HLookup(�����l, �����͈�, �s�ԍ�, �����̌^)
End Function
Public Function �悤��(���t�Z�� As Variant, ���1����3 As Variant) As Variant
    �悤�� = Application.WorksheetFunction.Weekday(���t�Z��, ���1����3)
End Function
Public Function �����Ƃ��߂�����(�Ώې��l�Z�� As Variant) As Variant
    �����Ƃ��߂����� = Application.WorksheetFunction.Even(�Ώې��l�Z��)
End Function
Public Function �����Ƃ��߂��(�Ώې��l�Z�� As Variant) As Variant
    �����Ƃ��߂�� = Application.WorksheetFunction.Odd(�Ώې��l�Z��)
End Function
Public Function �����Ƃ���������������(�Ώې��l�Z�� As Variant) As Variant
    �����Ƃ��������������� = Application.WorksheetFunction.Even(�Ώې��l�Z��)
End Function
Public Function �����Ƃ�������������(�Ώې��l�Z�� As Variant) As Variant
    �����Ƃ������������� = Application.WorksheetFunction.Odd(�Ώې��l�Z��)
End Function
Public Function ��������(�����͈� As Variant, ��r�l As Variant, ���ϔ͈� As Variant) As Variant
    �������� = Application.WorksheetFunction.AverageIf(�����͈�, ��r�l, ���ϔ͈�)
End Function
Public Function ����������łȂ�(�ΏۃZ�� As Variant) As Variant
    ����������łȂ� = Application.WorksheetFunction.IsNonText(�ΏۃZ��)
End Function
Public Function ����������(�ΏۃZ�� As Variant) As Variant
    ���������� = Application.WorksheetFunction.IsText(�ΏۃZ��)
End Function
Public Function �������l(�ΏۃZ�� As Variant) As Variant
    �������l = Application.WorksheetFunction.IsNumber(�ΏۃZ��)
End Function
Public Function �������v(�����͈� As Variant, ��r�l As Variant, ���v�͈� As Variant) As Variant
    �������v = Application.WorksheetFunction.SumIf(�����͈�, ��r�l, ���v�͈�)
End Function
Public Function ��������(�����͈� As Variant, ��r�l As Variant) As Variant
    �������� = Application.WorksheetFunction.CountIf(�����͈�, ��r�l)
End Function
Public Function ��������(�ΏۃZ�� As Variant) As Variant
    �������� = Application.WorksheetFunction.IsEven(�ΏۃZ��)
End Function
Public Function ������(�ΏۃZ�� As Variant) As Variant
    ������ = IsEmpty(�ΏۃZ��)
End Function
Public Function �����(�ΏۃZ�� As Variant) As Variant
    ����� = Application.WorksheetFunction.IsOdd(�ΏۃZ��)
End Function
Public Function ������Ȃ���(������ As Variant) As Variant
    ������Ȃ��� = Len(������)
End Function
Public Function ����������łȂ�(�ΏۃZ�� As Variant) As Variant
    ����������łȂ� = Application.WorksheetFunction.IsNonText(�ΏۃZ��)
End Function
Public Function �����������(�ΏۃZ�� As Variant) As Variant
    ����������� = Application.WorksheetFunction.IsText(�ΏۃZ��)
End Function
Public Function �����ւ�����(�����͈� As Variant, ��r�l As Variant, ���ϔ͈� As Variant) As Variant
    �����ւ����� = Application.WorksheetFunction.AVERGEIF(�����͈�, ��r�l, ���ϔ͈�)
End Function
Public Function �����̂��Ƃ��������(�ΏۃZ�� As Variant) As Variant
    �����̂��Ƃ�������� = Application.WorksheetFunction.IsNA(�ΏۃZ��)
End Function
Public Function ����������(�u���ΏۃZ�� As Variant, �u���Ώە����� As Variant, �u���㕶���� As Variant) As Variant
    ���������� = Application.WorksheetFunction.Substitute(�u���ΏۃZ��, �u���Ώە�����, �u���㕶����)
End Function
Public Function ����������(�ΏۃZ�� As Variant) As Variant
    ���������� = Application.WorksheetFunction.IsNumber(�ΏۃZ��)
End Function
Public Function ������������(�����͈� As Variant, ��r�l As Variant, ���v�͈� As Variant) As Variant
    ������������ = Application.WorksheetFunction.SumIf(�����͈�, ��r�l, ���v�͈�)
End Function
Public Function �������񂷂�(�����͈� As Variant, ��r�l As Variant) As Variant
    �������񂷂� = Application.WorksheetFunction.CountIf(�����͈�, ��r�l)
End Function
Public Function ���������͂�(�ΏۃZ�� As Variant) As Variant
    ���������͂� = IsEmpty(�ΏۃZ��)
End Function
Public Function ������������(�ΏۃZ�� As Variant) As Variant
    ������������ = Application.WorksheetFunction.IsEven(�ΏۃZ��)
End Function
Public Function ����������(�ΏۃZ�� As Variant) As Variant
    ���������� = Application.WorksheetFunction.IsOdd(�ΏۃZ��)
End Function
Public Function �����G���[(�ΏۃZ�� As Variant) As Variant
    �����G���[ = Application.WorksheetFunction.IsError(�ΏۃZ��)
End Function
Public Function ����NA(�ΏۃZ�� As Variant) As Variant
    ����NA = Application.WorksheetFunction.IsNA(�ΏۃZ��)
End Function
Public Function ����(������ As Variant, �^�l As Variant, �U�l As Variant) As Variant
    ���� = IIf(������, �^�l, �U�l)
End Function
Public Function �݂��������(������ As Variant, �E����̕����� As Variant) As Variant
    �݂�������� = Right(������, �E����̕�����)
End Function
Public Function �܂���(�_������1 As Variant, �_������2 As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    �܂��� = Application.WorksheetFunction.Or(�_������1, �_������2, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function �ׂ���(���̐��Z�� As Variant, �ׂ��搔�Z�� As Variant) As Variant
    �ׂ��� = Application.WorksheetFunction.Power(���̐��Z��, �ׂ��搔�Z��)
End Function
Public Function �ւ��ق�����(���l�Z�� As Variant) As Variant
    �ւ��ق����� = Sqr(���l�Z��)
End Function
Public Function �ւ�����(���ϔ͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    �ւ����� = Application.WorksheetFunction.Average(���ϔ͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function �Ԃ񂳂�(�Z���͈� As Variant) As Variant
    �Ԃ񂳂� = Application.WorksheetFunction.VarP(�Z���͈�)
End Function
Public Function �ӂ�(�����Z�� As Variant) As Variant
    �ӂ� = Minute(�����Z��)
End Function
Public Function �ӂ肪�ȕ\��(�Ώە����Z�� As Variant) As Variant
    �ӂ肪�ȕ\�� = Application.WorksheetFunction.Phonetic(�Ώە����Z��)
End Function
Public Function �ӂ肪��(�Ώە����Z�� As Variant) As Variant
    �ӂ肪�� = Application.WorksheetFunction.Phonetic(�Ώە����Z��)
End Function
Public Function �ӂւ�Ԃ񂳂�(�Z���͈� As Variant) As Variant
    �ӂւ�Ԃ񂳂� = Application.WorksheetFunction.Var(�Z���͈�)
End Function
Public Function �ӂւ�Ђ傤�����ւ�(�Z���͈� As Variant) As Variant
    �ӂւ�Ђ傤�����ւ� = Application.WorksheetFunction.StDev(�Z���͈�)
End Function
Public Function �Ђ傤�����ւ�(�Z���͈� As Variant) As Variant
    �Ђ傤�����ւ� = Application.WorksheetFunction.StDevP(�Z���͈�)
End Function
Public Function �т傤(�����Z�� As Variant) As Variant
    �т傤 = Second(�����Z��)
End Function
Public Function �ЂÂ��ւ񂩂�(�N As Variant, �� As Variant, �� As Variant) As Variant
    �ЂÂ��ւ񂩂� = DateSerial(�N, ��, ��)
End Function
Public Function �Ђ���������(������ As Variant, ������̕����� As Variant) As Variant
    �Ђ��������� = Left(������, ������̕�����)
End Function
Public Function ��(���t�Z�� As Variant) As Variant
    �� = Day(���t�Z��)
End Function
Public Function �΂��������肷��(���l As Variant, �{��� As Variant) As Variant
    �΂��������肷�� = Application.WorksheetFunction.Floor(���l, �{���)
End Function
Public Function �΂��������肠��(���l As Variant, �{��� As Variant) As Variant
    �΂��������肠�� = Application.WorksheetFunction.Ceiling(���l, �{���)
End Function
Public Function �˂�(���t�Z�� As Variant) As Variant
    �˂� = Year(���t�Z��)
End Function
Public Function �Ȃ񂵂イ�߂̂ЂÂ�(���t As Variant, �t���O1�܂���2 As Variant) As Variant
    �Ȃ񂵂イ�߂̂ЂÂ� = Application.WorksheetFunction.WeekNum(���t, �t���O1�܂���2)
End Function
Public Function ���Â���[��ւ񂳂������ւ񂳂������̂��񂫂�Ԃ�(���� As Variant, ���߂���͉̂����ڂ� As Variant, �ԍό��� As Variant, �ؓ����z As Variant, �Ō�Ɏc����z As Variant) As Variant
    ���Â���[��ւ񂳂������ւ񂳂������̂��񂫂�Ԃ� = Application.WorksheetFunction.PPmt(����, ���߂���͉̂����ڂ�, �ԍό���, �ؓ����z, �Ō�Ɏc����z)
End Function
Public Function ���Â���[��ւ񂳂������̂����Ԃ�(���� As Variant, ���߂���͉̂����ڂ� As Variant, �ԍό��� As Variant, �ؓ��z As Variant, �Ō�Ɏc����z As Variant) As Variant
    ���Â���[��ւ񂳂������̂����Ԃ� = Application.WorksheetFunction.IPmt(����, ���߂���͉̂����ڂ�, �ԍό���, �ؓ��z, �Ō�Ɏc����z)
End Function
Public Function ���Â���[��ւ񂳂�����(���� As Variant, �ԍό��� As Variant, �ؓ��z As Variant, �Ō�Ɏc����z As Variant) As Variant
    ���Â���[��ւ񂳂����� = Application.WorksheetFunction.Pmt(����, �ԍό���, �ؓ��z, �Ō�Ɏc����z)
End Function
Public Function ���Â��݂��Ă��傿���͂炢���݂���(���� As Variant, �ϗ����� As Variant, �������z As Variant, �ړI�̐ϗ��z As Variant) As Variant
    ���Â��݂��Ă��傿���͂炢���݂��� = Application.WorksheetFunction.Pmt(����, �ϗ�����, �������z, �ړI�̐ϗ��z)
End Function
Public Function ���イ������(�Z���͈� As Variant) As Variant
    ���イ������ = Application.WorksheetFunction.Median(�Z���͈�)
End Function
Public Function ���������ق�����Ȃ�΂��(�Z���͈� As Variant, ���� As Variant) As Variant
    ���������ق�����Ȃ�΂�� = Application.WorksheetFunction.Small(�Z���͈�, ����)
End Function
Public Function �^���W�F���g(���l�Z�� As Variant) As Variant
    �^���W�F���g = Tan(���l�Z��)
End Function
Public Function ���ĂЂ傤�т�(�����l As Variant, �����͈� As Variant, ��ԍ� As Variant, �����̌^ As Variant) As Variant
    ���ĂЂ傤�т� = Application.WorksheetFunction.VLookup(�����l, �����͈�, ��ԍ�, �����̌^)
End Function
Public Function ��������(���̐��l As Variant) As Variant
    �������� = Log(���̐��l)
End Function
Public Function ���񂩂����͂񂩂���(�Ώە����Z�� As Variant) As Variant
    ���񂩂����͂񂩂��� = Application.WorksheetFunction.Asc(�Ώە����Z��)
End Function
Public Function �Z������(�����͈� As Variant) As Variant
    �Z������ = Application.WorksheetFunction.CountA(�����͈�)
End Function
Public Function ���邯�񂷂�(�����͈� As Variant) As Variant
    ���邯�񂷂� = Application.WorksheetFunction.CountA(�����͈�)
End Function
Public Function ����������(���l�Z�� As Variant) As Variant
    ���������� = Abs(���l�Z��)
End Function
Public Function �����������񂾂�(�J�n�l As Variant, �I���l As Variant) As Variant
    �����������񂾂� = Application.WorksheetFunction.RandBetween(�J�n�l, �I���l)
End Function
Public Function ����������[�܂�������(�ΏۃZ�� As Variant) As Variant
    ����������[�܂������� = Application.WorksheetFunction.Roman(�ΏۃZ��)
End Function
Public Function �����������̂��܂�(�J�n�� As Variant, �� As Variant) As Variant
    �����������̂��܂� = Application.WorksheetFunction.EoMonth(�J�n��, ��)
End Function
Public Function ����������(�J�n�� As Variant, �� As Variant) As Variant
    ���������� = Application.WorksheetFunction.EDate(�J�n��, ��)
End Function
Public Function ���傤�悤��������(���̐��l As Variant) As Variant
    ���傤�悤�������� = Application.WorksheetFunction.Log10(���̐��l)
End Function
Public Function ���傤(�Ώې��l�Z�� As Variant, ���鐔 As Variant) As Variant
    ���傤 = Application.WorksheetFunction.Quotient(�Ώې��l�Z��, ���鐔)
End Function
Public Function �����(���ʒ����Z�� As Variant, �Z���͈� As Variant, �����t���O As Variant) As Variant
    ����� = Application.WorksheetFunction.Rank(���ʒ����Z��, �Z���͈�, �����t���O)
End Function
Public Function �����񂽂������̂Ă��ׂ̂����傤(�ׂ��ƂȂ鐔 As Variant) As Variant
    �����񂽂������̂Ă��ׂ̂����傤 = Exp(�ׂ��ƂȂ鐔)
End Function
Public Function �����񂽂�����(���̐��l As Variant) As Variant
    �����񂽂����� = Application.WorksheetFunction.Ln(���̐��l)
End Function
Public Function �����Ⴒ�ɂイ(���l As Variant, �l�̌ܓ����錅�� As Variant) As Variant
    �����Ⴒ�ɂイ = Application.WorksheetFunction.Round(���l, �l�̌ܓ����錅��)
End Function
Public Function ������ւ񂩂�(�� As Variant, �� As Variant, �b As Variant) As Variant
    ������ւ񂩂� = TimeSerial(��, ��, �b)
End Function
Public Function ������̂�(��r�P�� As Variant, ���t�Z��1 As Variant, ���t�Z��2 As Variant) As Variant
    ������̂� = DateDiff(��r�P��, ���t�Z��1, ���t�Z��2)
End Function
Public Function ��(�����Z�� As Variant) As Variant
    �� = Hour(�����Z��)
End Function
Public Function ���䂤�����͂���������(�Ώە����Z�� As Variant) As Variant
    ���䂤�����͂��������� = Application.WorksheetFunction.Trim(�Ώە����Z��)
End Function
Public Function �T�C��(���l�Z�� As Variant) As Variant
    �T�C�� = Sin(���l�Z��)
End Function
Public Function �����Ђ�(�Z���͈� As Variant) As Variant
    �����Ђ� = Application.WorksheetFunction.Mode(�Z���͈�)
End Function
Public Function �������������΂�����(���l�͈� As Variant) As Variant
    �������������΂����� = Application.WorksheetFunction.Gcd(���l�͈�)
End Function
Public Function ��������(�����͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    �������� = Application.WorksheetFunction.max(�����͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function �������傤�����΂�����(���l�͈� As Variant) As Variant
    �������傤�����΂����� = Application.WorksheetFunction.Lcm(���l�͈�)
End Function
Public Function �������傤(�����͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    �������傤 = Application.WorksheetFunction.Min(�����͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function ������(�Ώە����Z�� As Variant) As Variant
    ������ = LCase(�Ώە����Z��)
End Function
Public Function �R�T�C��(���l�Z�� As Variant) As Variant
    �R�T�C�� = Cos(���l�Z��)
End Function
Public Function ��������(���v�͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    �������� = Application.WorksheetFunction.Sum(���v�͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function ���񂷂�(�J�E���g�͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    ���񂷂� = Application.WorksheetFunction.Count(�J�E���g�͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function ����(���t�Z�� As Variant) As Variant
    ���� = Month(���t�Z��)
End Function
Public Function ���肩�����Ђ傤��(�Ώە����� As Variant, �J��Ԃ��� As Variant) As Variant
    ���肩�����Ђ傤�� = Application.WorksheetFunction.Rept(�Ώە�����, �J��Ԃ���)
End Function
Public Function ���肷��2(���l As Variant, �؂�̂Ă錅��� As Variant) As Variant
    ���肷��2 = Application.WorksheetFunction.RoundDown(���l, �؂�̂Ă錅���)
End Function
Public Function ���肷��(���l As Variant) As Variant
    ���肷�� = Int(���l)
End Function
Public Function ���肠��(���l As Variant, �؂�グ�錅�� As Variant) As Variant
    ���肠�� = Application.WorksheetFunction.RoundUp(���l, �؂�グ�錅��)
End Function
Public Function ����(�_������1 As Variant, �_������2 As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    ���� = Application.WorksheetFunction.And(�_������1, �_������2, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function ������(���W�A���Z�� As Variant) As Variant
    ������ = Application.WorksheetFunction.Degrees(���W�A���Z��)
End Function
Public Function �������傤(���̐��l As Variant) As Variant
    �������傤 = Application.WorksheetFunction.Fact(���̐��l)
End Function
Public Function ��������(�Ώە����Z�� As Variant) As Variant
    �������� = UCase(�Ώە����Z��)
End Function
Public Function ���������ق�����Ȃ�΂��(�Z���͈� As Variant, ���� As Variant) As Variant
    ���������ق�����Ȃ�΂�� = Application.WorksheetFunction.Large(�Z���͈�, ����)
End Function
Public Function ���񂵂イ���() As Variant
    ���񂵂イ��� = Application.WorksheetFunction.Pi
End Function
Public Function �������񂲂̂���Ƃ�����������������(�p�P����܂ރZ�� As Variant) As Variant
    �������񂲂̂���Ƃ����������������� = Application.WorksheetFunction.Proper(�p�P����܂ރZ��)
End Function
Public Function �������傤�тɂ�����(�J�n�� As Variant, �I���� As Variant, �Փ����������Z���͈� As Variant) As Variant
    �������傤�тɂ����� = Application.WorksheetFunction.NetworkDays(�J�n��, �I����, �Փ����������Z���͈�)
End Function
Public Function �������傤��(�J�n�� As Variant, ���� As Variant, �Փ��̓��t���������Z���͈� As Variant) As Variant
    �������傤�� = Application.WorksheetFunction.WorkDay(�J�n��, ����, �Փ��̓��t���������Z���͈�)
End Function
Public Function �C���f�b�N�X(�����͈� As Variant, �s�ԍ� As Variant, ��ԍ� As Variant) As Variant
    �C���f�b�N�X = Application.WorksheetFunction.index(�����͈�, �s�ԍ�, ��ԍ�)
End Function
Public Function ����() As Variant
    ���� = Now
End Function
Public Function ������(�����l As Variant, �����͈� As Variant, �ƍ��̎�� As Variant) As Variant
    ������ = Application.WorksheetFunction.Match(�����l, �����͈�, �ƍ��̎��)
End Function
Public Function �������������(������ As Variant, �擪�����ԍ� As Variant, �����o�������� As Variant) As Variant
    ������������� = Mid(������, �擪�����ԍ�, �����o��������)
End Function
Public Function �A�[�N�^���W�F���g(x���W As Variant, y���W As Variant) As Variant
    �A�[�N�^���W�F���g = Application.WorksheetFunction.Atan2(x���W, y���W)
End Function
Public Function �A�[�N�T�C��(���̃T�C���̐��l As Variant) As Variant
    �A�[�N�T�C�� = Application.WorksheetFunction.Asin(���̃T�C���̐��l)
End Function
Public Function �A�[�N�R�T�C��(���̃R�T�C���̐��l As Variant) As Variant
    �A�[�N�R�T�C�� = Application.WorksheetFunction.Acos(���̃R�T�C���̐��l)
End Function
Public Function nPr���ʂ�(�����Z�� As Variant, ���o�����Z�� As Variant) As Variant
    nPr���ʂ� = Application.WorksheetFunction.Permut(�����Z��, ���o�����Z��)
End Function

Public Function nCr���ʂ�(�����Z�� As Variant, ���o�����Z�� As Variant) As Variant
    nCr���ʂ� = Application.WorksheetFunction.Combin(�����Z��, ���o�����Z��)
End Function

Public Function ���t��(���t�V���A�� As Variant) As Date
    ���t�� = CDate(���t�V���A��)
End Function

Public Function ����(Optional �� As Variant = 255, Optional �� As Variant = 255, Optional �� As Variant = 255) As Variant
    ���� = RGB(��, ��, ��)
End Function

Public Function �F(Optional �� As Variant = 255, Optional �� As Variant = 255, Optional �� As Variant = 255) As Variant
    �F = RGB(��, ��, ��)
End Function

Public Function �F�C���f�b�N�X����RGB�F�֕ϊ�(idx As �J���[�C���f�b�N�X�p�^�[��) As Variant
    Select Case idx
    Case 1
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 0, 0)
    Case 2
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 255, 255)
    Case 3
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 0, 0)
    Case 4
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 255, 0)
    Case 5
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 0, 255)
    Case 6
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 255, 0)
    Case 7
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 0, 255)
    Case 8
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 255, 255)
    Case 9
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(128, 0, 0)
    Case 10
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 128, 0)
    Case 11
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 0, 128)
    Case 12
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(128, 128, 0)
    Case 13
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(128, 0, 128)
    Case 14
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 128, 128)
    Case 15
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(192, 192, 192)
    Case 16
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(128, 128, 128)
    Case 17
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(153, 153, 255)
    Case 18
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(153, 51, 102)
    Case 19
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 255, 204)
    Case 20
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(204, 255, 255)
    Case 21
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(102, 0, 102)
    Case 22
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 128, 128)
    Case 23
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 102, 204)
    Case 24
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(204, 204, 255)
    Case 25
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 0, 128)
    Case 26
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 0, 255)
    Case 27
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 255, 0)
    Case 28
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 255, 255)
    Case 29
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(128, 0, 128)
    Case 30
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(128, 0, 0)
    Case 31
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 128, 128)
    Case 32
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 0, 255)
    Case 33
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 204, 255)
    Case 34
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(204, 255, 255)
    Case 35
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(204, 255, 204)
    Case 36
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 255, 153)
    Case 37
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(153, 204, 255)
    Case 38
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 153, 204)
    Case 39
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(204, 153, 255)
    Case 40
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 204, 153)
    Case 41
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(51, 102, 255)
    Case 42
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(51, 204, 204)
    Case 43
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(153, 204, 0)
    Case 44
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 204, 0)
    Case 45
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 153, 0)
    Case 46
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 102, 0)
    Case 47
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(102, 102, 153)
    Case 48
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(150, 150, 150)
    Case 49
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 51, 102)
    Case 50
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(51, 153, 102)
    Case 51
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 51, 0)
    Case 52
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(51, 51, 0)
    Case 53
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(153, 51, 0)
    Case 54
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(153, 51, 102)
    Case 55
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(51, 51, 153)
    Case 56
        �F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(51, 51, 51)
    End Select

End Function

Public Function �F�̎O���F���擾(�F As Long, ByRef �� As Long, ByRef �� As Long, ByRef �� As Long)
    �� = �F Mod 256
    �� = Int(�F / 256) Mod 256
    �� = Int(�F / 256 / 256)
End Function

Public Function ����()
    ���� = Int(��())
End Function

Public Function �����̓��t()
    �����̓��t = Trim(���t��(����()))
End Function

Public Function ���̓��t()
    ���̓��t = ���t��(��())
End Function







