Attribute VB_Name = "My_lib"

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Mylib
'   update: 2004.11.2  T�W���֐��@�ύX���֖߂��i�����P�ցj�@�@s.f
'   update: 2004.9.26  T�W���֐��@�ύX�@�i�����Q�ցj�@s.f
'   update: 2002.6.28  s.f. public sub cal_pid �ǉ�
'   update: 2002.6.20 D/A�t���X�P�[���ύX(10V for 400kgf)
'   update: 2002.6.17 D/A�t���X�P�[���ύX02.6.17
'   update: 2005.11. 6 s.f   �I�[�o�[�t���[�΍�@�@long,double�֏����ւ� r_z!(),s_drive,setcm1
'   update: 2005.11.22 s.f   Melec C-870 counter����o�O�C���@�R���y�A�J�E���^�l�Z�b�g���@�������]�@�@setcm1
'   update: 2005.11.23 s.f   rstcm1 tsuika
'   update: 2005.11.26 s.f   �萔�́@����
'   update: 2005.12.23 s.f   longdata �v�Z�@1�s�@���@3�s
'   update: 2006. 5. 9 s.f    ppos = ppos & " r_z"
'   update: 2006. 5.14 s.f �@r_pres()�́@DoEvents �@ for�̊O�ֈړ��@s.f  ���̂���������
'�@�@�@�@�@�@�@�@�@�@�@�@�@�@���ׂĔ����Ɓ@LS_TC�@�v���O�����\������iLS_SC�́@OK)�f
'   update: 2006. 5.23 s.f �@cal_pid �ύX
'   update: 2006. 7.12 s.f �@my_lib �́@r_z!()�@w1,w2,w3 long �� integer
'
'

''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function r_z!()
   Dim LongData As Long
'   Dim Longr_z As Long
'   Dim w1, w2, w3  As Long
   Dim w1, w2, w3  As Integer    ' 2006.7.12 s.f.OverFlow�@�΍�
   If BrdFlg <> "ON" Then Exit Function
  '-------------------------- Z�ʒu�ǂݎ��
   Ack = MPL_IRDrive(hDev, MplData, MplResult)   '���݈ʒu�`�c�c�q�d�r�r�̕\��
   w1 = MplData.MPL_Data(1)
   w2 = MplData.MPL_Data(2)
   w3 = MplData.MPL_Data(3)
      ppos = Left(ppos, 17) & " (r_z)"
   LongData = (w1 * idc65536)           '2005.11. 6 s.f  2005.12.23
   LongData = LongData + (w2 * idc256)  '2005.11. 6 s.f  2005.12.23
   LongData = LongData + w3             '2005.11. 6 s.f  2005.12.23
   If LongData > idc8388607 Then LongData = LongData - idc16777216
   r_z = -LongData / gRev2Disp
   '
   'If r_z > 0.1 Then OrgOFF      '���_LED�@off  2002.10.9 KYOCERA
   '
End Function

Public Function r_pres!()
Dim i%, l%
Dim sumdt!
Dim dt!(0 To 7)
Dim adFlg As Long
  ppos = Left(ppos, 22) & " (r_pre)"
  sumdt = 0
'  DoEvents              '�@2006.5.14�@�ړ�  2006.5.18 �폜
  For l = 1 To 10
    AdRead dt(), adFlg
'    DoEvents                 ' 2006.5.14 for�̊O�ֈړ��@s.f  ���̂���������
    'sumdt = sumdt + (dt(2) - 2.07667) * 223.8   '�׏d 66kg�ōZ��
    sumdt = sumdt + dt(2) * 50#   'D/A�t���X�P�[���ύX02.7.31(10V for 500kgf)
    'sumdt = sumdt + dt(2) * 40  'D/A�t���X�P�[���ύX02.6.20(10V for 400kgf)
    'sumdt = sumdt + dt(2) * 15  'D/A�t���X�P�[���ύX02.6.17
    'sumdt = sumdt + dt(2) * 10  '�׏d **kg�ōZ��
  Next l
  r_pres = sumdt / 10# - 0#   '����
End Function

Public Sub s_drive(az!, v!)
Dim k_puls As Long, hspd As Long
Dim sb As Double
Dim i%, sts%
Dim idt1 As Long, idt2 As Long, idt3 As Long
Dim ihd As Long
Dim sn  As Long
Dim pos, azd As Double

'2002.10.9 KYOCERA
  sts = PCTrnsChk
  If sts = 1 Then
    MsgBox "�r�p�������I�@�^�]���s�s�\", vbCritical + vbOKOnly, "�v���I�ُ�"
    End
  End If

'--------------- ���x�̐ݒ�
  hspd = v * gRev2Disp / 60
'  If hspd > 400000 Then hspd = 400000  '02.5.11.sf
  If hspd > 800000 Then hspd = 800000
  If hspd < 77 Then hspd = 77
  
  Call MplDataSet(hspd, MplData)      '�h�m�b�q�d�l�d�m�s�`�k �h�m�c�d�w �c�q�h�u�d �b�n�l�l�`�m�c
  Ack = MPL_IWDrive(hDev, &H8, MplData, MplResult)
  
'--------------- �p���X���̎Z�o
  azd = az
  pos = r_z()
  k_puls = (azd - pos) * gRev2Disp + ddc05
  'If k_puls > 0 Then sn = 1 Else sn = -1
  'idt1 = Int(k_puls * sn / idc65536)
  'idt2 = Int((k_puls * sn - idt1 * idc65536) / idc256)
  'idt3 = k_puls * sn - idt1 * idc65536 - idt2 * idc256
'--------------- �C���N�������g����
  Ready_Wait    'while((inp(AX_STS)&1)!=0);
  'Data = idt1: Ack = MPL_BWDriveData1(hDev, Data, MplResult)   '
  'Data = idt2: Ack = MPL_BWDriveData2(hDev, Data, MplResult)   '
  'Data = idt3: Ack = MPL_BWDriveData3(hDev, Data, MplResult)   '
  'cmd = &H14: Ack = MPL_BWDriveCommand(hDev, cmd, MplResult)   '
  Call MplDataSet(k_puls, MplData)                    '�h�m�b�q�d�l�d�m�s�`�k �h�m�c�d�w �c�q�h�u�d �b�n�l�l�`�m�c
  Ack = MPL_IWDrive(hDev, &H14, MplData, MplResult)
End Sub
Public Sub rstcm1()
Dim zclear!
  zclear = -200#
  setcm1 zclear
End Sub
Public Sub setcm1(az!)
Dim k_puls As Long
Dim idt1, idt2, idt3, sn As Long
Dim i%
Dim azd As Double
'--------------- ���B�p���X���Z
  sn = 1
  azd = -az          ' ���̂�������Ȃ����@�u�|�v�Ő��퓮��@�@2005.11.22�@��.��.
  k_puls = azd * gRev2Disp + ddc05
'  idt1 = Int(k_puls * sn / idc65536)�@�@�@�@�@�@�@�@�@�@�@�@�f�@2005.11.22�@�@MPL_IWCounter�@�R�}���h�֏��ւ�
'  idt2 = Int((k_puls * sn - idt1 * idc65536) / idc256)
'  idt3 = k_puls * sn - idt1 * idc65536 - idt2 * idc256
'--------------- �R���p���[�^�@�P�ݒ�
  Ready_Wait    'while((inp(AX_STS)&1)!=0);
'  Data = idt1: Ack = MPL_BWCounterData1(hDev, Data, MplResult)   '
'  Data = idt2: Ack = MPL_BWCounterData2(hDev, Data, MplResult)   '
'  Data = idt3: Ack = MPL_BWCounterData3(hDev, Data, MplResult)   '
'  Cmd = &H1: Ack = MPL_BWCounterCommand(hDev, Cmd, MplResult)
   Call MplDataSet(k_puls, MplData)                    '�h�m�b�q�d�l�d�m�s�`�k �h�m�c�d�w �c�q�h�u�d �b�n�l�l�`�m�c
   Ack = MPL_IWCounter(hDev, &H1, MplData, MplResult)
End Sub
Public Sub Counter0()
Dim k_puls As Long
Dim i%, idt1!, idt2!, idt3!, sn%
'--------------- �J�E���^�O
  Ready_Wait    'while((inp(AX_STS)&1)!=0);
  Data = 0: Ack = MPL_BWCounterData1(hDev, Data, MplResult)   '
  Data = 0: Ack = MPL_BWCounterData2(hDev, Data, MplResult)   '
  Data = 0: Ack = MPL_BWCounterData3(hDev, Data, MplResult)   '
  Cmd = 0: Ack = MPL_BWCounterCommand(hDev, Cmd, MplResult)
End Sub
Public Sub cal_pid(m_sa!, m_p!, m_lim!)
'  float  m_sa,     /* �ݒ舳�� */
'         m_p,      /* �ݒ�o�l */
'         m_lim;    /* �ݒ胊�~�b�g�l */
Dim i%, ch%
'Dim i%, nout%, ch%, v!    nout,v ��Global�錾�� 2004.3.12
Dim pa!, per!       '/* float�i�P���x���������_�^)*/
  ppos = ppos + "csub"
  pa = r_pres()     '/* ���� */

'  If ((pa > 500#) Or (pa < -100#)) Then  '/* 500�j���ȏ�Ŕ���~ */
  If ((pa > 500#) Or (pa < -200#)) Then  '/* 500�j���ȏ�Ŕ���~ */  2012.7.1 henkou
'  If pa > m_sa + 200# Then '/* �w�舳�� + 200�j���ȏ�Ŕ���~ */
  hijyou                  ' 2006.5.23  -100�ȉ��@�ǉ�
  gemgmsg = gemgmsg + "cal_pid" + Format(pa, "0.0")   '2010.7.6 '2010.5.19 s.f.
    Exit Sub
  End If

'/* �o�h�c���Z */
  ppos = ppos + "1"
  per = 5# * (m_sa - pa) * Abs(m_sa - pa) / (m_p * m_p)
  If per > m_lim Then per = m_lim
  If per < (-1# * m_lim) Then per = -1# * m_lim     ' 2006.5.23 #�ǉ�
  'nout = Int(40.95 * per) + &H800
  ppos = ppos + "2"
  nout = &H800 - Int(4.095 * per / 4#)
  'nout = &H800 - Int(40.95 * per)
  ch = 1
'  v = 10# * (Int(4.095 * per / 4#) / 2048#)   ' 2005.11.26
  ppos = ppos + "3"
  DaOut ch, Hex(nout)
  'DaVoltOut ch, V
  'outp(ADPORT,(nout%256));
  'outp(ADPORT+1,0x20|(nout/idc256));
  
End Sub
Public Function T_keisu_cset!(t0cs!, tccs!)       ' 05.11.26�@s.f.�@�@overflow �΍� �u�I�v����
' /*  �V�ݒ艷�x�����x�W�����ݒ艷�x�@�@�́@�v�Z
' /* t00=�@�ݒ艷�x
' /* tc=�@���x�W��
'  Dim t0cs!, tccs!, abs0!
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  �v�Z�����@�P�@�@��Η�x����́@���
  Dim abs0!
   abs0 = -273#
'
   T_keisu_cset = (t0cs - abs0) * tccs + abs0
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  �v�Z�����@�Q�@�@���x�W���l�������@�V�t�g
'
'  Dim kijyun!, sa!
'
'  kijyun = 1#
'
'   T_keisu_cset = t0cs + (tccs - kijyun) * 100
'
End Function
Public Function T_keisu_cread!(t0cr!, tccr!)    ' 05.11.26�@s.f.�@�@overflow �΍� �u�I�v����
' /*  �V���݉��x�����݉��x/���x�W���@�@�́@�v�Z
' /* t00=�@�ݒ艷�x
' /* tc=�@���x�W��
'  Dim t0cr!, tccr!, abs0!
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  �v�Z�����@�P�@�@��Η�x����́@���
  Dim abs0!
'
   abs0 = -273#
'
   T_keisu_cread = (t0cr - abs0) / tccr + abs0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  �v�Z�����@�Q�@�@���x�W���l�������@�V�t�g
'
'  Dim kijyun!, sa!
'  kijyun = 1#
'
'    T_keisu_cread = t0cr - (tccr - kijyun) * 100
'
End Function

