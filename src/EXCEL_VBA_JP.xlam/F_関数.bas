Attribute VB_Name = "F_�֐�"
Option Explicit

Public Function F_�J�X�^��1(���v�͈� As Variant) As Variant
    F_�J�X�^��1 = F_�؂�̂�2(F_���v(���v�͈�), 2)
End Function
Public Function F_�J�X�^��2(�ΏۃZ�� As Variant) As Variant
    F_�J�X�^��2 = F_��(�ΏۃZ��, 12)
End Function
'���w�̃A�[�N�R�T�C���iarccos�j��x�ŕԂ��֐��ł��B
Public Function F_�A�[�N�R�T�C���x(cos�l As Variant) As Variant
    F_�A�[�N�R�T�C���x = Application.WorksheetFunction.Degrees(Application.WorksheetFunction.Acos(cos�l))
End Function

'���w�̃A�[�N�T�C���iarcsin�j��x�ŕԂ��֐��ł��B
Public Function F_�A�[�N�T�C���x(sin�l As Variant) As Variant
    F_�A�[�N�T�C���x = Application.WorksheetFunction.Degrees(Application.WorksheetFunction.Asin(sin�l))
End Function

'���w�̃A�[�N�^���W�F���g�iarctan�j���A�x�ŕԂ��֐��ł��B
Public Function F_�A�[�N�^���W�F���g�x(tan�l As Variant) As Variant
    F_�A�[�N�^���W�F���g�x = Application.WorksheetFunction.Degrees(Atn(tan�l))
End Function

'���w�̃R�T�C���icos�j��x����������֐��ł��B
Public Function F_�R�T�C���x(�x As Variant) As Variant
    F_�R�T�C���x = Cos(Application.WorksheetFunction.Radians(�x))
End Function

'���w�̃T�C���isin�j��x����������֐��ł��B
Public Function F_�T�C���x(�x As Variant) As Variant
    F_�T�C���x = Sin(Application.WorksheetFunction.Radians(�x))
End Function

'���w�̃^���W�F���g�x�itan�j��x����������֐��ł��B
Public Function F_�^���W�F���g�x(�x As Variant) As Variant
    F_�^���W�F���g�x = Tan(Application.WorksheetFunction.Radians(�x))
End Function

'2�i����10�i���ɕϊ����܂��B
Public Function F_��i������\�i��(��i�� As Variant) As Variant

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

    F_��i������\�i�� = �\�i���v�Z�p

End Function

' n �� m �̎��̗]������߂܂��B
Public Function F_�]��(�����鐔n As Variant, ���鐔m As Variant) As Variant
    F_�]�� = �����鐔n Mod ���鐔m
End Function

'16�i����10�i���ɕϊ����܂�
Public Function F_�\�Z�i������\�i��(�\�Z�i�� As Variant) As Variant
    F_�\�Z�i������\�i�� = Val("&H" & �\�Z�i��)
End Function

'10�i����2�i���ɕϊ����܂�
Public Function F_�\�i�������i��(�\�i�� As Variant, Optional ���� As Long = 8) As String
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

    F_�\�i�������i�� = Format(��i���v�Z�p, padding)
End Function

'10�i����16�i���ɕϊ����܂�
Public Function F_�\�i������\�Z�i��(�\�i�� As Variant, Optional ���� As Long = 4) As Variant
    F_�\�i������\�Z�i�� = F_�\�Z�i���p�f�B���O(Hex(�\�i��), "0", ����)
End Function
'�@�\�F�w�蕶�����ߊ֐�
'�����Fstr�@�F�ϊ��O�̕�����
'�@�@�@chr  �F���߂镶��(�P�����ڂ̂ݎg�p)
'�@�@�@digit�F����
'�ߒl�F�w�蕶�����ߌ�̕�����
Private Function F_�\�Z�i���p�f�B���O(ByVal str As String, _
                     ByVal char As String, _
                     ByVal digit As Long) As String
  Dim tmp As String
  tmp = str
  If Len(str) < digit And Len(char) > 0 Then
    tmp = Right(String(digit, char) & str, digit)
  End If
  F_�\�Z�i���p�f�B���O = tmp
End Function

'���K�\���̒u���p�^�[����������w�肵�āA���K�\���u�����܂��B
Public Function F_���K�\���u��(�����Ώ� As Variant, �u���p�^�[�������� As Variant, �u����̕����� As Variant, Optional �啶������������ As Boolean = False, Optional �ŏ��̈�v���̂ݒu�� As Boolean = False)
    r_RegExp.Pattern = �u���p�^�[��������
    r_RegExp.IgnoreCase = �啶������������
    r_RegExp.Global = Not �ŏ��̈�v���̂ݒu��
    If (IsObject(�����Ώ�)) Then
        F_���K�\���u�� = r_RegExp.Replace(�����Ώ�.Value2, �u����̕�����)
    Else
        F_���K�\���u�� = r_RegExp.Replace(�����Ώ�, �u����̕�����)
    End If
End Function


Public Function F_�j��(���t�Z�� As Variant, ���1����3 As Variant) As Variant
    F_�j�� = Application.WorksheetFunction.Weekday(���t�Z��, ���1����3)
End Function
Public Function F_������(���l�Z�� As Variant) As Variant
    F_������ = Sqr(���l�Z��)
End Function
Public Function F_����(���ϔ͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    F_���� = Application.WorksheetFunction.Average(���ϔ͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function F_������(������ As Variant) As Variant
    F_������ = Len(������)
End Function
Public Function F_�����u��(�u���ΏۃZ�� As Variant, �u���Ώە����� As Variant, �u���㕶���� As Variant) As Variant
    F_�����u�� = Application.WorksheetFunction.Substitute(�u���ΏۃZ��, �u���Ώە�����, �u���㕶����)
End Function
Public Function F_���U(�Z���͈� As Variant) As Variant
    F_���U = Application.WorksheetFunction.VarP(�Z���͈�)
End Function
Public Function F_��(�����Z�� As Variant) As Variant
    F_�� = Minute(�����Z��)
End Function
Public Function F_�s�Ε��U(�Z���͈� As Variant) As Variant
    F_�s�Ε��U = Application.WorksheetFunction.Var(�Z���͈�)
End Function
Public Function F_�s�ΕW���΍�(�Z���͈� As Variant) As Variant
    F_�s�ΕW���΍� = Application.WorksheetFunction.StDev(�Z���͈�)
End Function
Public Function F_�b(�����Z�� As Variant) As Variant
    F_�b = Second(�����Z��)
End Function
Public Function F_�W���΍�(�Z���͈� As Variant) As Variant
    F_�W���΍� = Application.WorksheetFunction.StDevP(�Z���͈�)
End Function
Public Function F_�{���؂�グ(���l As Variant, �{��� As Variant) As Variant
    F_�{���؂�グ = Application.WorksheetFunction.Ceiling(���l, �{���)
End Function
Public Function F_�{���؂�̂�(���l As Variant, �{��� As Variant) As Variant
    F_�{���؂�̂� = Application.WorksheetFunction.Floor(���l, �{���)
End Function
Public Function F_�N(���t�Z�� As Variant) As Variant
    F_�N = Year(���t�Z��)
End Function
Public Function F_���t�ϊ�(�N As Variant, �� As Variant, �� As Variant) As Variant
    F_���t�ϊ� = DateSerial(�N, ��, ��)
End Function
Public Function F_���t�̍�(��r�P�� As Variant, ���t�Z��1 As Variant, ���t�Z��2 As Variant) As Variant
    F_���t�̍� = DateDiff(��r�P��, ���t�Z��1, ���t�Z��2)
End Function
Public Function F_��(���t�Z�� As Variant) As Variant
    F_�� = Day(���t�Z��)
End Function
Public Function F_�����l(�Z���͈� As Variant) As Variant
    F_�����l = Application.WorksheetFunction.Median(�Z���͈�)
End Function
Public Function F_�傫�������牽�Ԗڂ��̒l(�Z���͈� As Variant, ���� As Variant) As Variant
    F_�傫�������牽�Ԗڂ��̒l = Application.WorksheetFunction.Large(�Z���͈�, ����)
End Function
Public Function F_�ΐ�(���̐��l As Variant) As Variant
    F_�ΐ� = Log(���̐��l)
End Function
Public Function F_�S�p�����𔼊p��(�Ώە����Z�� As Variant) As Variant
    F_�S�p�����𔼊p�� = Application.WorksheetFunction.Asc(�Ώە����Z��)
End Function
Public Function F_��Βl(���l�Z�� As Variant) As Variant
    F_��Βl = Abs(���l�Z��)
End Function
Public Function F_�؂�グ(���l As Variant, �؂�グ�錅�� As Variant) As Variant
    F_�؂�グ = Application.WorksheetFunction.RoundUp(���l, �؂�グ�錅��)
End Function
Public Function F_�؂�̂�2(���l As Variant, �؂�̂Ă錅��� As Variant) As Variant
    F_�؂�̂�2 = Application.WorksheetFunction.RoundDown(���l, �؂�̂Ă錅���)
End Function
Public Function F_�؂�̂�(���l As Variant) As Variant
    F_�؂�̂� = Int(���l)
End Function
Public Function F_���l�ԃ����_��(�J�n�l As Variant, �I���l As Variant) As Variant
    F_���l�ԃ����_�� = Application.WorksheetFunction.RandBetween(�J�n�l, �I���l)
End Function
Public Function F_���������[�}������(�ΏۃZ�� As Variant) As Variant
    F_���������[�}������ = Application.WorksheetFunction.Roman(�ΏۃZ��)
End Function
Public Function F_��������̌���(�J�n�� As Variant, �� As Variant) As Variant
    F_��������̌��� = Application.WorksheetFunction.EoMonth(�J�n��, ��)
End Function
Public Function F_��������(�J�n�� As Variant, �� As Variant) As Variant
    F_�������� = Application.WorksheetFunction.EDate(�J�n��, ��)
End Function
Public Function F_��p�ΐ�(���̐��l As Variant) As Variant
    F_��p�ΐ� = Application.WorksheetFunction.Log10(���̐��l)
End Function
Public Function F_�����������牽�Ԗڂ��̒l(�Z���͈� As Variant, ���� As Variant) As Variant
    F_�����������牽�Ԗڂ��̒l = Application.WorksheetFunction.Small(�Z���͈�, ����)
End Function
Public Function F_��(�Ώې��l�Z�� As Variant, ���鐔 As Variant) As Variant
    F_�� = Application.WorksheetFunction.Quotient(�Ώې��l�Z��, ���鐔)
End Function
Public Function F_����(���ʒ����Z�� As Variant, �Z���͈� As Variant, �����t���O As Variant) As Variant
    F_���� = Application.WorksheetFunction.Rank(���ʒ����Z��, �Z���͈�, �����t���O)
End Function
Public Function F_�c�\��(�����l As Variant, �����͈� As Variant, ��ԍ� As Variant, Optional �I�v�V����1 As Variant) As Variant
    F_�c�\�� = Application.WorksheetFunction.VLookup(�����l, �����͈�, ��ԍ�, �I�v�V����1)
End Function
Public Function F_���R�ΐ��̒�e�ׂ̂���(�ׂ��ƂȂ鐔 As Variant) As Variant
    F_���R�ΐ��̒�e�ׂ̂��� = Exp(�ׂ��ƂȂ鐔)
End Function
Public Function F_���R�ΐ�(���̐��l As Variant) As Variant
    F_���R�ΐ� = Application.WorksheetFunction.Ln(���̐��l)
End Function
Public Function F_���ԕϊ�(�� As Variant, �� As Variant, �b As Variant) As Variant
    F_���ԕϊ� = TimeSerial(��, ��, �b)
End Function
Public Function F_��(�����Z�� As Variant) As Variant
    F_�� = Hour(�����Z��)
End Function
Public Function F_�l�̌ܓ�(���l As Variant, �l�̌ܓ����錅�� As Variant) As Variant
    F_�l�̌ܓ� = Application.WorksheetFunction.Round(���l, �l�̌ܓ����錅��)
End Function
Public Function F_�ŕp�l(�Z���͈� As Variant) As Variant
    F_�ŕp�l = Application.WorksheetFunction.Mode(�Z���͈�)
End Function
Public Function F_�ő����(���l�͈� As Variant) As Variant
    F_�ő���� = Application.WorksheetFunction.Gcd(���l�͈�)
End Function
Public Function F_�ő�(�����͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    F_�ő� = Application.WorksheetFunction.max(�����͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function F_�ŏ����{��(���l�͈� As Variant) As Variant
    F_�ŏ����{�� = Application.WorksheetFunction.Lcm(���l�͈�)
End Function
Public Function F_�ŏ�(�����͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    F_�ŏ� = Application.WorksheetFunction.Min(�����͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function F_��������(������ As Variant, ������̕����� As Variant) As Variant
    F_�������� = Left(������, ������̕�����)
End Function
Public Function F_���E�󔒕����폜(�Ώە����Z�� As Variant) As Variant
    F_���E�󔒕����폜 = Application.WorksheetFunction.Trim(�Ώە����Z��)
End Function
Public Function F_��() As Variant
    F_�� = Now
End Function
Public Function F_���v(���v�͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    F_���v = Application.WorksheetFunction.Sum(���v�͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function F_����(�J�E���g�͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    F_���� = Application.WorksheetFunction.Count(�J�E���g�͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function F_���X�ϗ����~�����z(���� As Variant, �ϗ����� As Variant, �������z As Variant, �ړI�̐ϗ��z As Variant) As Variant
    F_���X�ϗ����~�����z = Application.WorksheetFunction.Pmt(����, �ϗ�����, �������z, �ړI�̐ϗ��z)
End Function
Public Function F_���X���[���ԍϊz���̌����ԍϊz(���� As Variant, ���߂���͉̂����ڂ� As Variant, �ԍό��� As Variant, �ؓ����z As Variant, �Ō�Ɏc����z As Variant) As Variant
    F_���X���[���ԍϊz���̌����ԍϊz = Application.WorksheetFunction.PPmt(����, ���߂���͉̂����ڂ�, �ԍό���, �ؓ����z, �Ō�Ɏc����z)
End Function
Public Function F_���X���[���ԍϊz���̋������z(���� As Variant, ���߂���͉̂����ڂ� As Variant, �ԍό��� As Variant, �ؓ��z As Variant, �Ō�Ɏc����z As Variant) As Variant
    F_���X���[���ԍϊz���̋������z = Application.WorksheetFunction.IPmt(����, ���߂���͉̂����ڂ�, �ԍό���, �ؓ��z, �Ō�Ɏc����z)
End Function
Public Function F_���X���[���ԍϊz(���� As Variant, �ԍό��� As Variant, �ؓ��z As Variant, �Ō�Ɏc����z As Variant) As Variant
    F_���X���[���ԍϊz = Application.WorksheetFunction.Pmt(����, �ԍό���, �ؓ��z, �Ō�Ɏc����z)
End Function
Public Function F_��(���t�Z�� As Variant) As Variant
    F_�� = Month(���t�Z��)
End Function
Public Function F_�J��Ԃ��\��(�Ώە����� As Variant, �J��Ԃ��� As Variant) As Variant
    F_�J��Ԃ��\�� = Application.WorksheetFunction.Rept(�Ώە�����, �J��Ԃ���)
End Function
Public Function F_�ԕ�����(������ As Variant, �擪�����ԍ� As Variant, �����o�������� As Variant) As Variant
    F_�ԕ����� = Mid(������, �擪�����ԍ�, �����o��������)
End Function
Public Function F_�p�x(���W�A���Z�� As Variant) As Variant
    F_�p�x = Application.WorksheetFunction.Degrees(���W�A���Z��)
End Function
Public Function F_�K��(���̐��l As Variant) As Variant
    F_�K�� = Application.WorksheetFunction.Fact(���̐��l)
End Function
Public Function F_�����ڂ̓��t(���t As Variant, �t���O1�܂���2 As Variant) As Variant
    F_�����ڂ̓��t = Application.WorksheetFunction.WeekNum(���t, �t���O1�܂���2)
End Function
Public Function F_���\��(�����l As Variant, �����͈� As Variant, �s�ԍ� As Variant, Optional �I�v�V����1 As Variant) As Variant
    F_���\�� = Application.WorksheetFunction.HLookup(�����l, �����͈�, �s�ԍ�, �I�v�V����1)
End Function
Public Function F_�~����() As Variant
    F_�~���� = Application.WorksheetFunction.Pi
End Function
Public Function F_�p�P��̐擪������啶����(�p�P����܂ރZ�� As Variant) As Variant
    F_�p�P��̐擪������啶���� = Application.WorksheetFunction.Proper(�p�P����܂ރZ��)
End Function
Public Function F_�p���啶����(�Ώە����Z�� As Variant) As Variant
    F_�p���啶���� = UCase(�Ώە����Z��)
End Function
Public Function F_�p����������(�Ώە����Z�� As Variant) As Variant
    F_�p���������� = LCase(�Ώە����Z��)
End Function
Public Function F_�c�Ɠ�����(�J�n�� As Variant, �I���� As Variant, �Փ����������Z���͈� As Variant) As Variant
    F_�c�Ɠ����� = Application.WorksheetFunction.NetworkDays(�J�n��, �I����, �Փ����������Z���͈�)
End Function
Public Function F_�c�Ɠ�(�J�n�� As Variant, ���� As Variant, �Փ��̓��t���������Z���͈� As Variant) As Variant
    F_�c�Ɠ� = Application.WorksheetFunction.WorkDay(�J�n��, ����, �Փ��̓��t���������Z���͈�)
End Function
Public Function F_�E������(������ As Variant, �E����̕����� As Variant) As Variant
    F_�E������ = Right(������, �E����̕�����)
End Function
Public Function F_��v(�����l As Variant, �����͈� As Variant, �ƍ��̎�� As Variant) As Variant
    F_��v = Application.WorksheetFunction.Match(�����l, �����͈�, �ƍ��̎��)
End Function
Public Function F_���[���ԍϊz�̗��q�������̗݌v�z(���� As Variant, ���[���_�񌎐� As Variant, �ؓ��z As Variant, ��n�� As Variant, ��m�� As Variant, ���������Ȃ�0���񕥂��Ȃ�1 As Variant) As Variant
    F_���[���ԍϊz�̗��q�������̗݌v�z = Application.WorksheetFunction.CumIPmt(����, ���[���_�񌎐�, �ؓ��z, ��n��, ��m��, ���������Ȃ�0���񕥂��Ȃ�1)
End Function
Public Function F_���[���ԍϊz�̌����������̗݌v�z(���� As Variant, ���[���_�񌎐� As Variant, �ؓ��z As Variant, ��n�� As Variant, ��m�� As Variant, ���������Ȃ�0���񕥂��Ȃ�1 As Variant) As Variant
    F_���[���ԍϊz�̌����������̗݌v�z = Application.WorksheetFunction.CumPrinc(����, ���[���_�񌎐�, �ؓ��z, ��n��, ��m��, ���������Ȃ�0���񕥂��Ȃ�1)
End Function
Public Function F_��[��ւ񂳂������̂肵�Ԃ�̂邯��������(���� As Variant, ���[���_�񌎐� As Variant, �ؓ��z As Variant, ��n�� As Variant, ��m�� As Variant, ���������Ȃ�0���񕥂��Ȃ�1 As Variant) As Variant
    F_��[��ւ񂳂������̂肵�Ԃ�̂邯�������� = Application.WorksheetFunction.CumIPmt(����, ���[���_�񌎐�, �ؓ��z, ��n��, ��m��, ���������Ȃ�0���񕥂��Ȃ�1)
End Function
Public Function F_��[��ւ񂳂������̂��񂫂�Ԃ�̂邢��������(���� As Variant, ���[���_�񌎐� As Variant, �ؓ��z As Variant, ��n�� As Variant, ��m�� As Variant, ���������Ȃ�0���񕥂��Ȃ�1 As Variant) As Variant
    F_��[��ւ񂳂������̂��񂫂�Ԃ�̂邢�������� = Application.WorksheetFunction.CumPrinc(����, ���[���_�񌎐�, �ؓ��z, ��n��, ��m��, ���������Ȃ�0���񕥂��Ȃ�1)
End Function
Public Function F_�����_��() As Variant
    F_�����_�� = Rnd
End Function
Public Function F_���W�A��(�p�x�Z�� As Variant) As Variant
    F_���W�A�� = Application.WorksheetFunction.Radians(�p�x�Z��)
End Function
Public Function F_�悱�Ђ傤�т�(�����l As Variant, �����͈� As Variant, �s�ԍ� As Variant, �����̌^ As Variant) As Variant
    F_�悱�Ђ傤�т� = Application.WorksheetFunction.HLookup(�����l, �����͈�, �s�ԍ�, �����̌^)
End Function
Public Function F_�悤��(���t�Z�� As Variant, ���1����3 As Variant) As Variant
    F_�悤�� = Application.WorksheetFunction.Weekday(���t�Z��, ���1����3)
End Function
Public Function F_�����Ƃ��߂�����(�Ώې��l�Z�� As Variant) As Variant
    F_�����Ƃ��߂����� = Application.WorksheetFunction.Even(�Ώې��l�Z��)
End Function
Public Function F_�����Ƃ��߂��(�Ώې��l�Z�� As Variant) As Variant
    F_�����Ƃ��߂�� = Application.WorksheetFunction.Odd(�Ώې��l�Z��)
End Function
Public Function F_�����Ƃ���������������(�Ώې��l�Z�� As Variant) As Variant
    F_�����Ƃ��������������� = Application.WorksheetFunction.Even(�Ώې��l�Z��)
End Function
Public Function F_�����Ƃ�������������(�Ώې��l�Z�� As Variant) As Variant
    F_�����Ƃ������������� = Application.WorksheetFunction.Odd(�Ώې��l�Z��)
End Function
Public Function F_��������(�����͈� As Variant, ��r�l As Variant, ���ϔ͈� As Variant) As Variant
    F_�������� = Application.WorksheetFunction.AverageIf(�����͈�, ��r�l, ���ϔ͈�)
End Function
Public Function F_����������łȂ�(�ΏۃZ�� As Variant) As Variant
    F_����������łȂ� = Application.WorksheetFunction.IsNonText(�ΏۃZ��)
End Function
Public Function F_����������(�ΏۃZ�� As Variant) As Variant
    F_���������� = Application.WorksheetFunction.IsText(�ΏۃZ��)
End Function
Public Function F_�������l(�ΏۃZ�� As Variant) As Variant
    F_�������l = Application.WorksheetFunction.IsNumber(�ΏۃZ��)
End Function
Public Function F_�������v(�����͈� As Variant, ��r�l As Variant, ���v�͈� As Variant) As Variant
    F_�������v = Application.WorksheetFunction.SumIf(�����͈�, ��r�l, ���v�͈�)
End Function
Public Function F_��������(�����͈� As Variant, ��r�l As Variant) As Variant
    F_�������� = Application.WorksheetFunction.CountIf(�����͈�, ��r�l)
End Function
Public Function F_��������(�ΏۃZ�� As Variant) As Variant
    F_�������� = Application.WorksheetFunction.IsEven(�ΏۃZ��)
End Function
Public Function F_������(�ΏۃZ�� As Variant) As Variant
    F_������ = IsEmpty(�ΏۃZ��)
End Function
Public Function F_�����(�ΏۃZ�� As Variant) As Variant
    F_����� = Application.WorksheetFunction.IsOdd(�ΏۃZ��)
End Function
Public Function F_������Ȃ���(������ As Variant) As Variant
    F_������Ȃ��� = Len(������)
End Function
Public Function F_����������łȂ�(�ΏۃZ�� As Variant) As Variant
    F_����������łȂ� = Application.WorksheetFunction.IsNonText(�ΏۃZ��)
End Function
Public Function F_�����������(�ΏۃZ�� As Variant) As Variant
    F_����������� = Application.WorksheetFunction.IsText(�ΏۃZ��)
End Function
Public Function F_�����ւ�����(�����͈� As Variant, ��r�l As Variant, ���ϔ͈� As Variant) As Variant
    F_�����ւ����� = Application.WorksheetFunction.AVERGEIF(�����͈�, ��r�l, ���ϔ͈�)
End Function
Public Function F_�����̂��Ƃ��������(�ΏۃZ�� As Variant) As Variant
    F_�����̂��Ƃ�������� = Application.WorksheetFunction.IsNA(�ΏۃZ��)
End Function
Public Function F_����������(�u���ΏۃZ�� As Variant, �u���Ώە����� As Variant, �u���㕶���� As Variant) As Variant
    F_���������� = Application.WorksheetFunction.Substitute(�u���ΏۃZ��, �u���Ώە�����, �u���㕶����)
End Function
Public Function F_����������(�ΏۃZ�� As Variant) As Variant
    F_���������� = Application.WorksheetFunction.IsNumber(�ΏۃZ��)
End Function
Public Function F_������������(�����͈� As Variant, ��r�l As Variant, ���v�͈� As Variant) As Variant
    F_������������ = Application.WorksheetFunction.SumIf(�����͈�, ��r�l, ���v�͈�)
End Function
Public Function F_�������񂷂�(�����͈� As Variant, ��r�l As Variant) As Variant
    F_�������񂷂� = Application.WorksheetFunction.CountIf(�����͈�, ��r�l)
End Function
Public Function F_���������͂�(�ΏۃZ�� As Variant) As Variant
    F_���������͂� = IsEmpty(�ΏۃZ��)
End Function
Public Function F_������������(�ΏۃZ�� As Variant) As Variant
    F_������������ = Application.WorksheetFunction.IsEven(�ΏۃZ��)
End Function
Public Function F_����������(�ΏۃZ�� As Variant) As Variant
    F_���������� = Application.WorksheetFunction.IsOdd(�ΏۃZ��)
End Function
Public Function F_�����G���[(�ΏۃZ�� As Variant) As Variant
    F_�����G���[ = Application.WorksheetFunction.IsError(�ΏۃZ��)
End Function
Public Function F_����NA(�ΏۃZ�� As Variant) As Variant
    F_����NA = Application.WorksheetFunction.IsNA(�ΏۃZ��)
End Function
Public Function F_����(������ As Variant, �^�l As Variant, �U�l As Variant) As Variant
    F_���� = IIf(������, �^�l, �U�l)
End Function
Public Function F_�݂��������(������ As Variant, �E����̕����� As Variant) As Variant
    F_�݂�������� = Right(������, �E����̕�����)
End Function
Public Function F_�܂���(�_������1 As Variant, �_������2 As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    F_�܂��� = Application.WorksheetFunction.Or(�_������1, �_������2, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function F_�ׂ���(���̐��Z�� As Variant, �ׂ��搔�Z�� As Variant) As Variant
    F_�ׂ��� = Application.WorksheetFunction.Power(���̐��Z��, �ׂ��搔�Z��)
End Function
Public Function F_�ւ��ق�����(���l�Z�� As Variant) As Variant
    F_�ւ��ق����� = Sqr(���l�Z��)
End Function
Public Function F_�ւ�����(���ϔ͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    F_�ւ����� = Application.WorksheetFunction.Average(���ϔ͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function F_�Ԃ񂳂�(�Z���͈� As Variant) As Variant
    F_�Ԃ񂳂� = Application.WorksheetFunction.VarP(�Z���͈�)
End Function
Public Function F_�ӂ�(�����Z�� As Variant) As Variant
    F_�ӂ� = Minute(�����Z��)
End Function
Public Function F_�ӂ肪�ȕ\��(�Ώە����Z�� As Variant) As Variant
    F_�ӂ肪�ȕ\�� = Application.WorksheetFunction.Phonetic(�Ώە����Z��)
End Function
Public Function F_�ӂ肪��(�Ώە����Z�� As Variant) As Variant
    F_�ӂ肪�� = Application.WorksheetFunction.Phonetic(�Ώە����Z��)
End Function
Public Function F_�ӂւ�Ԃ񂳂�(�Z���͈� As Variant) As Variant
    F_�ӂւ�Ԃ񂳂� = Application.WorksheetFunction.Var(�Z���͈�)
End Function
Public Function F_�ӂւ�Ђ傤�����ւ�(�Z���͈� As Variant) As Variant
    F_�ӂւ�Ђ傤�����ւ� = Application.WorksheetFunction.StDev(�Z���͈�)
End Function
Public Function F_�Ђ傤�����ւ�(�Z���͈� As Variant) As Variant
    F_�Ђ傤�����ւ� = Application.WorksheetFunction.StDevP(�Z���͈�)
End Function
Public Function F_�т傤(�����Z�� As Variant) As Variant
    F_�т傤 = Second(�����Z��)
End Function
Public Function F_�ЂÂ��ւ񂩂�(�N As Variant, �� As Variant, �� As Variant) As Variant
    F_�ЂÂ��ւ񂩂� = DateSerial(�N, ��, ��)
End Function
Public Function F_�Ђ���������(������ As Variant, ������̕����� As Variant) As Variant
    F_�Ђ��������� = Left(������, ������̕�����)
End Function
Public Function F_��(���t�Z�� As Variant) As Variant
    F_�� = Day(���t�Z��)
End Function
Public Function F_�΂��������肷��(���l As Variant, �{��� As Variant) As Variant
    F_�΂��������肷�� = Application.WorksheetFunction.Floor(���l, �{���)
End Function
Public Function F_�΂��������肠��(���l As Variant, �{��� As Variant) As Variant
    F_�΂��������肠�� = Application.WorksheetFunction.Ceiling(���l, �{���)
End Function
Public Function F_�˂�(���t�Z�� As Variant) As Variant
    F_�˂� = Year(���t�Z��)
End Function
Public Function F_�Ȃ񂵂イ�߂̂ЂÂ�(���t As Variant, �t���O1�܂���2 As Variant) As Variant
    F_�Ȃ񂵂イ�߂̂ЂÂ� = Application.WorksheetFunction.WeekNum(���t, �t���O1�܂���2)
End Function
Public Function F_���Â���[��ւ񂳂������ւ񂳂������̂��񂫂�Ԃ�(���� As Variant, ���߂���͉̂����ڂ� As Variant, �ԍό��� As Variant, �ؓ����z As Variant, �Ō�Ɏc����z As Variant) As Variant
    F_���Â���[��ւ񂳂������ւ񂳂������̂��񂫂�Ԃ� = Application.WorksheetFunction.PPmt(����, ���߂���͉̂����ڂ�, �ԍό���, �ؓ����z, �Ō�Ɏc����z)
End Function
Public Function F_���Â���[��ւ񂳂������̂����Ԃ�(���� As Variant, ���߂���͉̂����ڂ� As Variant, �ԍό��� As Variant, �ؓ��z As Variant, �Ō�Ɏc����z As Variant) As Variant
    F_���Â���[��ւ񂳂������̂����Ԃ� = Application.WorksheetFunction.IPmt(����, ���߂���͉̂����ڂ�, �ԍό���, �ؓ��z, �Ō�Ɏc����z)
End Function
Public Function F_���Â���[��ւ񂳂�����(���� As Variant, �ԍό��� As Variant, �ؓ��z As Variant, �Ō�Ɏc����z As Variant) As Variant
    F_���Â���[��ւ񂳂����� = Application.WorksheetFunction.Pmt(����, �ԍό���, �ؓ��z, �Ō�Ɏc����z)
End Function
Public Function F_���Â��݂��Ă��傿���͂炢���݂���(���� As Variant, �ϗ����� As Variant, �������z As Variant, �ړI�̐ϗ��z As Variant) As Variant
    F_���Â��݂��Ă��傿���͂炢���݂��� = Application.WorksheetFunction.Pmt(����, �ϗ�����, �������z, �ړI�̐ϗ��z)
End Function
Public Function F_���イ������(�Z���͈� As Variant) As Variant
    F_���イ������ = Application.WorksheetFunction.Median(�Z���͈�)
End Function
Public Function F_���������ق�����Ȃ�΂��(�Z���͈� As Variant, ���� As Variant) As Variant
    F_���������ق�����Ȃ�΂�� = Application.WorksheetFunction.Small(�Z���͈�, ����)
End Function
Public Function F_�^���W�F���g(���l�Z�� As Variant) As Variant
    F_�^���W�F���g = Tan(���l�Z��)
End Function
Public Function F_���ĂЂ傤�т�(�����l As Variant, �����͈� As Variant, ��ԍ� As Variant, �����̌^ As Variant) As Variant
    F_���ĂЂ傤�т� = Application.WorksheetFunction.VLookup(�����l, �����͈�, ��ԍ�, �����̌^)
End Function
Public Function F_��������(���̐��l As Variant) As Variant
    F_�������� = Log(���̐��l)
End Function
Public Function F_���񂩂����͂񂩂���(�Ώە����Z�� As Variant) As Variant
    F_���񂩂����͂񂩂��� = Application.WorksheetFunction.Asc(�Ώە����Z��)
End Function
Public Function F_�Z������(�����͈� As Variant) As Variant
    F_�Z������ = Application.WorksheetFunction.CountA(�����͈�)
End Function
Public Function F_���邯�񂷂�(�����͈� As Variant) As Variant
    F_���邯�񂷂� = Application.WorksheetFunction.CountA(�����͈�)
End Function
Public Function F_����������(���l�Z�� As Variant) As Variant
    F_���������� = Abs(���l�Z��)
End Function
Public Function F_�����������񂾂�(�J�n�l As Variant, �I���l As Variant) As Variant
    F_�����������񂾂� = Application.WorksheetFunction.RandBetween(�J�n�l, �I���l)
End Function
Public Function F_����������[�܂�������(�ΏۃZ�� As Variant) As Variant
    F_����������[�܂������� = Application.WorksheetFunction.Roman(�ΏۃZ��)
End Function
Public Function F_�����������̂��܂�(�J�n�� As Variant, �� As Variant) As Variant
    F_�����������̂��܂� = Application.WorksheetFunction.EoMonth(�J�n��, ��)
End Function
Public Function F_����������(�J�n�� As Variant, �� As Variant) As Variant
    F_���������� = Application.WorksheetFunction.EDate(�J�n��, ��)
End Function
Public Function F_���傤�悤��������(���̐��l As Variant) As Variant
    F_���傤�悤�������� = Application.WorksheetFunction.Log10(���̐��l)
End Function
Public Function F_���傤(�Ώې��l�Z�� As Variant, ���鐔 As Variant) As Variant
    F_���傤 = Application.WorksheetFunction.Quotient(�Ώې��l�Z��, ���鐔)
End Function
Public Function F_�����(���ʒ����Z�� As Variant, �Z���͈� As Variant, �����t���O As Variant) As Variant
    F_����� = Application.WorksheetFunction.Rank(���ʒ����Z��, �Z���͈�, �����t���O)
End Function
Public Function F_�����񂽂������̂Ă��ׂ̂����傤(�ׂ��ƂȂ鐔 As Variant) As Variant
    F_�����񂽂������̂Ă��ׂ̂����傤 = Exp(�ׂ��ƂȂ鐔)
End Function
Public Function F_�����񂽂�����(���̐��l As Variant) As Variant
    F_�����񂽂����� = Application.WorksheetFunction.Ln(���̐��l)
End Function
Public Function F_�����Ⴒ�ɂイ(���l As Variant, �l�̌ܓ����錅�� As Variant) As Variant
    F_�����Ⴒ�ɂイ = Application.WorksheetFunction.Round(���l, �l�̌ܓ����錅��)
End Function
Public Function F_������ւ񂩂�(�� As Variant, �� As Variant, �b As Variant) As Variant
    F_������ւ񂩂� = TimeSerial(��, ��, �b)
End Function
Public Function F_������̂�(��r�P�� As Variant, ���t�Z��1 As Variant, ���t�Z��2 As Variant) As Variant
    F_������̂� = DateDiff(��r�P��, ���t�Z��1, ���t�Z��2)
End Function
Public Function F_��(�����Z�� As Variant) As Variant
    F_�� = Hour(�����Z��)
End Function
Public Function F_���䂤�����͂���������(�Ώە����Z�� As Variant) As Variant
    F_���䂤�����͂��������� = Application.WorksheetFunction.Trim(�Ώە����Z��)
End Function
Public Function F_�T�C��(���l�Z�� As Variant) As Variant
    F_�T�C�� = Sin(���l�Z��)
End Function
Public Function F_�����Ђ�(�Z���͈� As Variant) As Variant
    F_�����Ђ� = Application.WorksheetFunction.Mode(�Z���͈�)
End Function
Public Function F_�������������΂�����(���l�͈� As Variant) As Variant
    F_�������������΂����� = Application.WorksheetFunction.Gcd(���l�͈�)
End Function
Public Function F_��������(�����͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    F_�������� = Application.WorksheetFunction.max(�����͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function F_�������傤�����΂�����(���l�͈� As Variant) As Variant
    F_�������傤�����΂����� = Application.WorksheetFunction.Lcm(���l�͈�)
End Function
Public Function F_�������傤(�����͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    F_�������傤 = Application.WorksheetFunction.Min(�����͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function F_������(�Ώە����Z�� As Variant) As Variant
    F_������ = LCase(�Ώە����Z��)
End Function
Public Function F_�R�T�C��(���l�Z�� As Variant) As Variant
    F_�R�T�C�� = Cos(���l�Z��)
End Function
Public Function F_��������(���v�͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    F_�������� = Application.WorksheetFunction.Sum(���v�͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function F_���񂷂�(�J�E���g�͈� As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    F_���񂷂� = Application.WorksheetFunction.Count(�J�E���g�͈�, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function F_����(���t�Z�� As Variant) As Variant
    F_���� = Month(���t�Z��)
End Function
Public Function F_���肩�����Ђ傤��(�Ώە����� As Variant, �J��Ԃ��� As Variant) As Variant
    F_���肩�����Ђ傤�� = Application.WorksheetFunction.Rept(�Ώە�����, �J��Ԃ���)
End Function
Public Function F_���肷��2(���l As Variant, �؂�̂Ă錅��� As Variant) As Variant
    F_���肷��2 = Application.WorksheetFunction.RoundDown(���l, �؂�̂Ă錅���)
End Function
Public Function F_���肷��(���l As Variant) As Variant
    F_���肷�� = Int(���l)
End Function
Public Function F_���肠��(���l As Variant, �؂�グ�錅�� As Variant) As Variant
    F_���肠�� = Application.WorksheetFunction.RoundUp(���l, �؂�グ�錅��)
End Function
Public Function F_����(�_������1 As Variant, �_������2 As Variant, Optional �I�v�V����1 As Variant, Optional �I�v�V����2 As Variant, Optional �I�v�V����3 As Variant, Optional �I�v�V����4 As Variant, Optional �I�v�V����5 As Variant) As Variant
    F_���� = Application.WorksheetFunction.And(�_������1, �_������2, �I�v�V����1, �I�v�V����2, �I�v�V����3, �I�v�V����4, �I�v�V����5)
End Function
Public Function F_������(���W�A���Z�� As Variant) As Variant
    F_������ = Application.WorksheetFunction.Degrees(���W�A���Z��)
End Function
Public Function F_�������傤(���̐��l As Variant) As Variant
    F_�������傤 = Application.WorksheetFunction.Fact(���̐��l)
End Function
Public Function F_��������(�Ώە����Z�� As Variant) As Variant
    F_�������� = UCase(�Ώە����Z��)
End Function
Public Function F_���������ق�����Ȃ�΂��(�Z���͈� As Variant, ���� As Variant) As Variant
    F_���������ق�����Ȃ�΂�� = Application.WorksheetFunction.Large(�Z���͈�, ����)
End Function
Public Function F_���񂵂イ���() As Variant
    F_���񂵂イ��� = Application.WorksheetFunction.Pi
End Function
Public Function F_�������񂲂̂���Ƃ�����������������(�p�P����܂ރZ�� As Variant) As Variant
    F_�������񂲂̂���Ƃ����������������� = Application.WorksheetFunction.Proper(�p�P����܂ރZ��)
End Function
Public Function F_�������傤�тɂ�����(�J�n�� As Variant, �I���� As Variant, �Փ����������Z���͈� As Variant) As Variant
    F_�������傤�тɂ����� = Application.WorksheetFunction.NetworkDays(�J�n��, �I����, �Փ����������Z���͈�)
End Function
Public Function F_�������傤��(�J�n�� As Variant, ���� As Variant, �Փ��̓��t���������Z���͈� As Variant) As Variant
    F_�������傤�� = Application.WorksheetFunction.WorkDay(�J�n��, ����, �Փ��̓��t���������Z���͈�)
End Function
Public Function F_�C���f�b�N�X(�����͈� As Variant, �s�ԍ� As Variant, ��ԍ� As Variant) As Variant
    F_�C���f�b�N�X = Application.WorksheetFunction.index(�����͈�, �s�ԍ�, ��ԍ�)
End Function
Public Function F_����() As Variant
    F_���� = Now
End Function
Public Function F_������(�����l As Variant, �����͈� As Variant, �ƍ��̎�� As Variant) As Variant
    F_������ = Application.WorksheetFunction.Match(�����l, �����͈�, �ƍ��̎��)
End Function
Public Function F_�������������(������ As Variant, �擪�����ԍ� As Variant, �����o�������� As Variant) As Variant
    F_������������� = Mid(������, �擪�����ԍ�, �����o��������)
End Function
Public Function F_�A�[�N�^���W�F���g(x���W As Variant, y���W As Variant) As Variant
    F_�A�[�N�^���W�F���g = Application.WorksheetFunction.Atan2(x���W, y���W)
End Function
Public Function F_�A�[�N�T�C��(���̃T�C���̐��l As Variant) As Variant
    F_�A�[�N�T�C�� = Application.WorksheetFunction.Asin(���̃T�C���̐��l)
End Function
Public Function F_�A�[�N�R�T�C��(���̃R�T�C���̐��l As Variant) As Variant
    F_�A�[�N�R�T�C�� = Application.WorksheetFunction.Acos(���̃R�T�C���̐��l)
End Function
Public Function F_nPr���ʂ�(�����Z�� As Variant, ���o�����Z�� As Variant) As Variant
    F_nPr���ʂ� = Application.WorksheetFunction.Permut(�����Z��, ���o�����Z��)
End Function

Public Function F_nCr���ʂ�(�����Z�� As Variant, ���o�����Z�� As Variant) As Variant
    F_nCr���ʂ� = Application.WorksheetFunction.Combin(�����Z��, ���o�����Z��)
End Function

Public Function F_���t��(���t�V���A�� As Variant) As Date
    F_���t�� = CDate(���t�V���A��)
End Function

Public Function F_����(Optional �� As Variant = 255, Optional �� As Variant = 255, Optional �� As Variant = 255) As Variant
    F_���� = RGB(��, ��, ��)
End Function

Public Function F_�F(Optional �� As Variant = 255, Optional �� As Variant = 255, Optional �� As Variant = 255) As Variant
    F_�F = RGB(��, ��, ��)
End Function

Public Function F_�F�C���f�b�N�X����RGB�F�֕ϊ�(idx As �J���[�C���f�b�N�X�p�^�[��) As Variant
    Select Case idx
    Case 1
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 0, 0)
    Case 2
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 255, 255)
    Case 3
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 0, 0)
    Case 4
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 255, 0)
    Case 5
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 0, 255)
    Case 6
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 255, 0)
    Case 7
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 0, 255)
    Case 8
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 255, 255)
    Case 9
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(128, 0, 0)
    Case 10
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 128, 0)
    Case 11
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 0, 128)
    Case 12
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(128, 128, 0)
    Case 13
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(128, 0, 128)
    Case 14
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 128, 128)
    Case 15
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(192, 192, 192)
    Case 16
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(128, 128, 128)
    Case 17
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(153, 153, 255)
    Case 18
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(153, 51, 102)
    Case 19
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 255, 204)
    Case 20
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(204, 255, 255)
    Case 21
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(102, 0, 102)
    Case 22
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 128, 128)
    Case 23
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 102, 204)
    Case 24
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(204, 204, 255)
    Case 25
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 0, 128)
    Case 26
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 0, 255)
    Case 27
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 255, 0)
    Case 28
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 255, 255)
    Case 29
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(128, 0, 128)
    Case 30
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(128, 0, 0)
    Case 31
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 128, 128)
    Case 32
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 0, 255)
    Case 33
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 204, 255)
    Case 34
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(204, 255, 255)
    Case 35
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(204, 255, 204)
    Case 36
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 255, 153)
    Case 37
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(153, 204, 255)
    Case 38
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 153, 204)
    Case 39
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(204, 153, 255)
    Case 40
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 204, 153)
    Case 41
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(51, 102, 255)
    Case 42
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(51, 204, 204)
    Case 43
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(153, 204, 0)
    Case 44
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 204, 0)
    Case 45
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 153, 0)
    Case 46
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(255, 102, 0)
    Case 47
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(102, 102, 153)
    Case 48
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(150, 150, 150)
    Case 49
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 51, 102)
    Case 50
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(51, 153, 102)
    Case 51
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(0, 51, 0)
    Case 52
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(51, 51, 0)
    Case 53
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(153, 51, 0)
    Case 54
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(153, 51, 102)
    Case 55
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(51, 51, 153)
    Case 56
        F_�F�C���f�b�N�X����RGB�F�֕ϊ� = RGB(51, 51, 51)
    End Select

End Function

Public Function F_�F�̎O���F���擾(�F As Long, ByRef �� As Long, ByRef �� As Long, ByRef �� As Long)
    �� = �F Mod 256
    �� = Int(�F / 256) Mod 256
    �� = Int(�F / 256 / 256)
End Function

Public Function F_����()
    F_���� = Int(F_��())
End Function

Public Function F_�����̓��t()
    F_�����̓��t = Trim(F_���t��(F_����()))
End Function

Public Function F_���̓��t()
    F_���̓��t = F_���t��(F_��())
End Function





