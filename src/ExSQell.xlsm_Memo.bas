'========================================================
' ExSQell.xlsm - Copyright (C) 2016 nmrmsys
' sqltools.xls - Copyright (C) 2005 M.Nomura
'========================================================
'
'changelog
' v0.0.1 �ݒ�V�[�g��{�\���ASQL���s���� ExecSql ST_Query ST_Plus
' v0.0.2 �V�X�e�����[�e�B���e�B�̊g�[ GetCfgVal GetConStr Prop�N���X
' v0.0.3 �����m�F���W�b�N�ǉ� ST_Query ExecSql�A�R�����g���� [GS]etCmtVal
' v0.0.4 �}�N���N�������̕ύX Aplication.OnKey�AGetCfgSheet
' v0.0.5 ���ʃV�[�g�쐬�܂��𐮗� GetNewSheet MakeList
' v0.0.6 ��^SQL�̃V�[�g�Ǘ��A�v���[�X�z���_�u�� GetLibSql SetSqlParam
' v0.0.7 ST_Which�̎��� �I�u�W�F�N�g���擾 GetDBObjectType/List/Info
' v0.0.8 GetSql�g�� ���r��SQL�擾�A�\���^�C�v���f�A�^�C�g���擾�ASELECT�⊮
' v0.0.9 �Ȗ͗l�������t�������Ŏ����A�t�B���^�i���ݎ����ȁX�ɂȂ�悤�ɂ���
' v0.1.0 DML/DDL/DCL���s ST_Query ExecSql
' v0.1.1 �����m�F�̂� or �T���v�����o ST_Query MakeList
' v0.1.2 �g���ꗗ���� ST_Which GetDBObjectList MakeList
' v0.1.3 �֗��ȃV���[�g�J�b�g�L�[�ǉ� �V�[�g�폜�A�ݒ�V�[�g�A�O�V�[�g
' v0.1.4 �g�����擾/�\�� ST_ExtInfo GetDBObjectExtInfo
' v0.1.5 �I�u�W�F�N�g���擾 INDEX, PACKAGE�ǉ� GetDBObjectInfo
' v0.1.6 �����e�i���X���[�h ST_Maintain(�r���܂�)�AMakeList �󔒗� �s�A��
' v0.1.7 ��^�r�p�k���j���[�\�� ST_Library(�r���܂�)
' v0.1.8 BugFix MakeList ���ږ��ł̒l�擾�������ɁA�������ƂƎ擾�ł��Ȃ���
' v0.1.9 �N�G�����s���̎Ȗ͗l�����̃Z�b�g��ݒ�Ő���ł���悤�ɂ���
' v0.2.0 �����e�i���X���[�h�̍X�VSQL�����A���s���������쐬
' v0.2.1 �N�G�����s���̌����A�b���̃R�����g�t����ݒ�Ő���ł���悤�ɂ���
' v0.2.2 �����e���[�h�X�V �X�e�[�^�X�o�[�\���A�������ōĒ��o�A�����Ē��o
' v0.2.3 �I�[�g�t�B���^�����S�����̃V���[�g�J�b�g��ǉ�
'          DisplayFormulas = true�iCtrl+Shift+@�j�Ő����������ł���Ƃ�E�E
' v0.2.4 ���s�L��SQL�����Z���s�ɃZ�b�g����֐��A�󔒃Z���ł� Ctrl+W�̖���
' v0.2.5 �V�[�g���͐ݒ荀�ڂ���Ȃ��萔�ɕύX�A�s�ԍ��\�� �l/����/��\����I��
' v0.2.6 �����e�V�[�g�ɃC�x���g�}�N����t�����Ď����� U or I �}�[�N���Z�b�g
' v0.2.7 ���o�����V�[�g�\���A���͌� Ctrl + Q �ŃN�G�����s
' v0.2.8 �V�K��ƃV�[�g�쐬�̃V���[�g�J�b�g�ǉ�
' v0.2.9 �\�����\�I�u�W�F�N�g�̃N�G�����s���͒��o�����V�[�g��\��
' v0.3.0 �����e�i���X�V�[�g�̃L�[���ڃ^�C�g���𑾎��\��
' v0.3.1 ��^SQL�|�b�v�A�b�v���j���[ ���s���ʕ\���A��^SQL�V�[�g�\��
' v0.3.2 ���o�����V�[�g���\���ɂ��ď����w����g���񂷃��[�h�ǉ�
' v0.3.3 �V�[�g��\���A��\���V�[�g�ꗗ�V���[�g�J�b�g�L�[�ǉ�
' v0.3.4 ���o�����̓��͂������ANULL�w��AIN�w��A�ǉ�SQL���̃R�����g�Ή�
' v0.3.5 �R�[�h�̐����AMakeList�Ńf�[�^���o���ɐi���󋵂��X�e�[�^�X�\��
' v0.3.6 ���C�u�����ɂ��C�ɓ���u�b�N��o�^���|�b�v�A�b�v���j���[�\��
' v0.3.7 GetSqlParams,GetSql SQL������&�u���ϐ�������Γ���DLG�\��
' v0.3.8 ST_Template,GetTemplateSql �e���v���[�gSQL���̐���
' v0.3.9 �I�u�W�F�N�g�ꗗ�\���Ń^�C�v/���̂��t�B���^
' v0.4.0 ���o�����V�[�g ���ړ��{�ꖼ�̂�\�����ɁA�e�[�u�����G�C���A�X�t��
' v0.4.1 �N�G�����ʃV�[�g�̃f�[�^�^�����F�ASQL*Plus�̃E�C���h�E�^�C�g���ύX
' v0.4.2 �f�[�^�\��t�����A�����ݒ�𐔒l��ȊO�ɂ��������񏑎����ݒ�
' v0.4.3 �I��͈͂�SQL�����s SELECT�̂� �����̊��ʂ����鎞�͑Ή����銇�ʂ܂�
'          �����̎��s�ΏۊO�u���b�N�̔���͓���̂Ŏ�肠�������u
' v0.4.4 ExecSql Scroll Lock���͎��sSQL���E�C���h�E�\��
' v0.4.5 �e�[�u���\�����̃V�[�g���̓v���t�B�b�N�X + �e�[�u�����ŏ㏑���쐬
'          �����e�i���X�V�[�g�����l�ɃV�[�g���̃v���t�B�b�N�X��萔��
' v0.4.6 �R�[�h����SQL���o�ASQL�̃R�[�h��
' v0.4.7 SQL�̐��` ��肠���� SqlFmt.dll�g�p�̋@�\��t���Ƃ��A���̂�������
' v0.4.8 �X�VSQL�̎��s�ݒ�Ɏ��s���Ȃ���t���A���f�胁�b�Z�[�W��\��������
' v0.4.9 ���sSQL�̃E�C���h�E�\���ɃN�G�������s����/���Ȃ��̃{�^����ǉ�
' v0.5.0 �N�G�����ʃV�[�g�� A1 or �󔒃Z���� Ctrl + Q ���͍Ē��o�������Ȃ�
' v0.5.1 �X�VSQL�� ;���s ������ꍇ�͕������ăo�b�`���s�A�I��͈͎w�����
' v0.5.2 �����t�����Ŗ����Ȗ͗l�A�񕝎��������Ƀf�[�^�s/��S�̂̐ݒ��ǉ�
' v0.5.3 �B���@�\�̒��o�����V�[�g�̃��b�N�A���ږ��̕\���ASQL�\�������J����
' v0.5.4 �P�Ƃ̍X�VSQL���s���o�O���Ă����̂��C���A���C�u����URL���j���[�ǉ�
' v0.5.5 �����ڂ̃f�[�^�^���肪���������̂� oo4o�̎��� .OraIDataType���g�p
' v0.5.6 �f�[�^�\��t������ESC�����ŏ������f�A�����e���͂O���`�F�b�N����
' v0.5.7 ���o���� �ǉ�SQL��AND�t�����C���A��s�}���A�������͂�OR������\��
' v0.5.8 OO4O�ł̃��R�[�h�Z�b�g�擾�A�X�VSQL���s�̔񓯊����s ESC�L�[�Œ��f
' v0.5.9 �X�VSQL���s���̉e�����R�[�h�����\�� �O���Ȃ�\�����Ȃ��悤�ɂ���
' v0.6.0 �I��͈͂�SQL�����擾���ĕҏW�E�C���h�E��\���A�C���������e���ēW�J
' v0.6.1 �ҏW�E�C���h�E�̋@�\�����A�󔒃Z���������\���A�d���`�F�b�N��̒ǉ�
' v0.6.2 ��`�ꗗ�\�����ɑ��݂���I�u�W�F�N�g�Ȃ玩���Ńt�B���^�I���݂̂���
' v0.6.3 �X�g�A�h�̍ăR���p�C���A�͈͑I���Ńo�b�`�ăR���p�C�����\
' v0.6.4 oo4o��CreateDynaset���̃I�v�V�����w�肪���낢��s���������̂ŏC��
' v0.6.5 �X�g�A�h�̃R���p�C���A�G���[�����\�[�X�s�F�t���A�R�����g�ŕ\��
' v0.6.6 �X�g�A�h�̎��s
' v0.6.7 ���C�u�����n�R�}���h �L�[�A�T�C���ύX�A�o�^SQL/�����N�ǉ�
' v0.6.8 �ڑ��I�u�W�F�N�g�̃v�[�����O�A�R�~�b�g���[�h/���j���[�̒ǉ�
' v0.6.9 AUTOTRACE��ݒ肵����Ԃł� Plus�N��
' v0.7.0 �R�~�b�g���j���[ �X�VSQL�����s���Ă��鎞�� * �����ɕ\��
' v0.7.1 Worksheet for Oracle Utility�𓋍�
' v0.7.2 ���o���� ���t�^NULL�w��s��A�����l���C���h�J�[�h�̏ꍇLIKE OR�W�J
' v0.7.3 ���C�u�����Ƀi�C�X�~�h����SQL��ǉ��A�N�G�����s�Œu���ϐ�����DLG�\��
' v0.7.4 �����m�FSQL����������ORDER BY���܂܂�Ă���ꍇ�͈ȍ~���폜
'
'todo
'  �X�g�A�h�̃p�����[�^���擾 oo4o:USER_ARGUMENTS �e���v���W�J
'  �X�g�A�h�̎��s DBMS_OUTPUT��GET_LINES�Ŏ擾���A���ʕ\��
'  �����e/���o�����V�[�g��SQL�������ɃV�[�g������e�[�u�������擾���Ȃ�
'  �I�u�W�F�N�g�ꗗ�\���� Like���Z�q�𗘗p�����I��\��
'  �g���ꗗ�\�����@�ύX MakeList���ɏ������s�A�s�ĕ`��
'  �g�����̎�舵�� Prop��Keys�v���p�e�B�g�p
'  GetDBObject* �ʃX�L�[�}/�c�a�Ή�
'  SQL*Plus�𗠂Ŏ��s���Č��ʂ��擾 ���s�v��A�g���[�X�A���̑��Ŏg�p
'  �R�~�b�g���[�h �ڑ��I�u�W�F�N�g�̊Ǘ� �R���T�o���[�h�Ŏ��� memo�Q��
'  �L�[���f�̃��W�b�N���듮�삵�ăG���[���ɖ������[�v�ɂȂ��Ă�
'    �}�N�����s���̒��f�L�[����/�G���[����̕��j�𐮗�����K�v����
'  SQL���s��񓯊����s�ɂ��ăL�[���f���\�ɂ���
'    oo4o CreateSQL ORASQL_NONBLK(&H4) �� OraSqlStmt���쐬�� NonBlockingState
'    ado  adAsyncExecute adAsyncFetch���g�p�� State���Ď� Cancel�Œ��f
'  SQL���`�G���W�� ReSQL
'  LogParser�̃��R�[�h�Z�b�g����舵����悤�ɂ���
'    COM���͌`���v���O�C���Ŋe��h�L�������g�`���������\
'
'�R�}���h�ꗗ(���������܂�)
'  Ctrl + Q SQL���s
'    SHIFT���� �����m�FDLG�\��
'    ALT���� �����m�F�A���o�������̓L�����Z��
'    �\�A�r���[��SELECT�����⊮�A���o�������́A���������m�F
'  Ctrl + M �f�[�^�����e�i���X
'    SHIFT���� �����m�FDLG�\��
'    ALT���� �����m�F�A���o�������̓L�����Z��
'    �\�A�r���[��SELECT�����⊮�A���o�������́A���������m�F
'  Ctrl + P SQL*PLUS�N��
'    SHIFT���� ���s�v��: ��
'    ALT���� �g���[�X: ��
'  Ctrl + W �I�u�W�F�N�g�ꗗ/��`(�\��`�A�\�[�X)
'    �I�u�W�F�N�g�����ł����ꍇ�͒�`�\��
'    SHIFT���� �ꗗ�\���L�����Z��
'    ALT���� �g���ꗗ���ڂ��\���ASHIFT�����ňꗗ�\���F��
'    �I���Z����1�Â����F��
'  Ctrl + E �g�������R�����g�ŕt��
'    SHIFT���� �o���[���\�� �`���[�g�I�u�W�F�N�g�̕�����������
'    ALT���� ��`���V�[�g: ��
'    �I���Z����1�Â����F��
'  Ctrl + T �e���v���[�g����
'    SHIFT���� SQL���`
'    ALT���� SQL������
'  Ctrl + L �֗��Ȉꗗ���|�b�v�A�b�v���j���[�I�����\��
'    SHIFT���� SQL�}��
'    ALT���� ��^SQL�V�[�g�\��
'  Ctrl + K �I�����C���}�j���A���F��
'  Ctrl+Shift+INS     �V�K��ƃV�[�g�쐬
'  Ctrl+Shift+DEL     ���݃V�[�g�̍폜
'  Ctrl+Shift+BS      ���O�̃V�[�g��\��
'  Ctrl+Shift+TAB     �ݒ�V�[�g��\��
'  Ctrl+Shift+ *      �t�B���^/���o�����̃N���A
'  Ctrl+Shift+ -      �V�[�g�̔�\��
'  Ctrl+Shift+ ^      ��\���V�[�g�ꗗ
'  Alt+Shift+�����L�[ ���݃Z���ƒl�̈Ⴄ�Z���Ɉړ��F��
'
'memo
'
'���ʁF
'  �ڑ��֘A
'    A1�Z���ɐڑ������񂾂��łȂ� �L�[:�l �`��������
'    �R�~�b�g���[�h �ڑ��I�u�W�F�N�g�̊Ǘ�
'      �R�l�N�V�����T�[�o ConServ�ڑ����[�h �����ăR���T�o���[�h
'      SQL���s������ MultiUse��1�X���b�h�v�[���� ActiveX.EXE�Ŏ���
'      �ڑ��L�[���w�肷�鎖�ɂ��قȂ�N���C�A���g������R�l�N�V����
'      �g�����U�N�V���������L�ł��� �^�X�N�g���C���猻�݂̐ڑ��ꗗ�A
'      �R�~�b�g/���[���o�b�N����A���sSQL�����̊m�F �}���`�X���b�h��
'      ����SQL*Plus����HTA�N���C�A���g SQL*Minus
'      http://www.int21.co.jp/pcdn/vb/noriolib/vbmag/9809/com/
'      http://www.int21.co.jp/pcdn/vb/noriolib/vbmag/9810/db_solu/
'      http://www.koalanet.ne.jp/~akiya/vbtaste/vbp/#else03
'      http://www2.plala.or.jp/k-world/vbasic/vbasic009.html
'      http://pooh3.dip.jp/vbvcjava/window/tasktray.html
'      http://www2.netf.org/freesoft.html#IPCATL
'  ��^SQL�Ǘ�
'    &�ϐ��擾�p GetSqlParams accept��prompt������ꍇ�͂�����擾
'    accept�͂���܂�Ӗ����������������u���b�N�̌��ʂ��擾�ł��Ȃ�
'      DBMS_OUTPUT.GET_LINE�Ŏ擾�ł��鎖�����������X�g�A�h���s��
'      oo4o�ł�邱�Ƃɂ����̂ł��܂�g����������
'    SQL�������R�[�h�����SQL�����o������Ƃ��ɕϐ�����&�����ēW�J
'  ���ʃV�[�g����
'    1 �s���Ƃ̐F�h���d���s���m�������t������
'      http://www.excel7.com/chotto16.htm
'      http://www.excel7.com/chotto17.htm
'      �O�s�Ɠ����ꍇ�͔�\�� �f�[�^�sA�� �����t�������� =A2=A1
'      �����e�i���X�V�[�g�s�ǉ����Ή��p�̐���No.��
'  �֗��@�\
'    ���Z���ʒu�ɃW�����v HYPERLINK�֐�
'    �d���s���m�����Z�b�g�A�V�[�g���m�̍�����r
'
'ST_Query�F
'  �P��SELECT���̎��s
'  DML�𔻕ʂ����s INSERT UPDATE DELETE MERGE
'  DDL�𔻕ʂ����s CREATE ALTER DROP RENAME TRUNCATE COMMENT
'  DCL�𔻕ʂ����s GRANT REVOKE
'  �I�u�W�F�N�g�������̏ꍇ�͎�����SQL���W�J
'    TABLE, VIEW SELECT���W�J�A���o�������̓V�[�g�\��
'      http://chibie2000.hp.infoseek.co.jp/excelvba/
'      ���[�U�[�_�C�A���O�̓O���b�h�\�����ł��Ȃ��̂Ŏ~��
'    PROCEDURE, FUNCTION EXEC���F��
'  �����e�i���X�V�[�g�𔻕ʂ�DB�X�V
'  SP�𔻕ʂ��������ͤ���s����ʕ\��:��
'  SP�𔻕ʂ�SHIFT���������ōăR���p�C��:��
'    ALTER PROCEDUR/FUNCTION objname COMPILE
'ST_Which�F
'  �I�u�W�F�N�g���̎擾
'    OO4O: �f�B�N�V���i�� or OraMetaData(key,index�擾�s��)
'      �V�m�j���̉���
'      �ʃX�L�[�}���̎擾 DBA_* owner�w��v
'      �ʃf�[�^�x�[�X���̎擾 OraMetaData or �ʐڑ�
'      �g�����̕\��
'      PUBLIC�������Ώۂɂ��邵�Ȃ�
'    ADO:  �f�B�N�V���i�� (�ڑ��悪oracle�̏ꍇ�̂�) or ADOX
'  Private Function GetDBObjectType(argObjName) As String
'    OO4O: �J�^���O�� *_OBJECTS
'    ADO:   ADOX�� Tables, Views, Procedures
'  Private Function GetDBObjectList(argObjName, argRs, argFlds, argOpts)
'    OO4O: �J�^���O�� *_OBJECTS�A�g������ *_*S �ƃW���C��
'    ADO:   ADOX�� Tables, Views, Procedures
'  Private Function GetDBObjectInfo(argObjName, argRs, argFlds, argOpts)
'     OO4O: �J�^���O�� *_*S
'     ADO:   ADOX�� Table(Columns, Indexes, Keys), View / Procedure(Command)
'  ��`�V�[�g TABLE, VIEW, INDEX
'    ��`�V�[�g�͐擪�Z�����I�u�W�F�N�g���Ńt�B���^�\
'    ���̤��
'    ���́A�s���A�}�[�N��A���ڒ�` or �R�[�h
'    �����t��������2�s�ڈȍ~�͋�
'  �\�[�X�V�[�g PROCEDURE, FUNCTION, PACKAGE
'  ���̑��V�[�g SEQENCE, TRIGGER, SNAPSHOT, DBLINK
'  �g����� �V�m�j������ ���� �s�� �T�C�Y
'  Ctrl+Alt+Shift+W �͒�`�\�����L�����Z��
'  �Z���͈͑I���� Ctrl+W �͈ꗗ�\���L�����Z��
'  DBA_* �� �f�B�N�V���i���e�[�u���̏�񂪎擾�ł��邩 ���Ȃ�
'ST_ExtInfo�F
'  �g�������R�����g�t��
'  �g�������o���[���\��
'  ��`������
'  �I���Z����1�Â���:   ��
'ST_Maintenance�F
'  �f�[�^�ҏW�V�[�g
'    A����\���A1�ɐڑ��椃e�[�u����L�[
'      ����ȍ~�ɂ�TAB��؂�̍X�V�L�[�l��ۑ�
'    B��� i: �ǉ� U: �X�V D: �폜 ���w��
'    ���[�N�V�[�g�C�x���g�ŃZ���̍X�V���擾
'  �X�V�L�[���ROWID
'  VBA����̃C�x���g�}�N���}�� http://www.mrexcel.com/archive/VBA/14890.html
'ST_Plus�F
'  �ڑ���������E�C���h�E�^�C�g���\��:
'    http://homepage2.nifty.com/sak/w_sak3/doc/sysbrd/vb_t02.htm
'  SHIFT�����Ŏ��s�v��:   ��
'  ALT�����Ńg���[�X�\��:   ��
'ST_Template�F
'  �\�A�r���[��SELECT/INSERT/UPDATE/DELETE/CREATE
'  SP�͎��s�p�̖����u���b�N�𐶐�
'  ���̑��̃I�u�W�F�N�g��DDL�𐶐�
'  SQL���` SqlFmt.dll�͂��܈�A���O�Ŏ���
'  SQL������/����
'ST_Library�F
'  �R�}���h�o�[�Ń|�b�v�A�b�v���j���[
'  ���[�U��Z�b�V������g�p��
'  ��^SQL�V�[�g��FLG��ON�̂��̂̂ݕ\�� �ꗗ Or �}��
'ST_Knowledge�F
'  http://www.oracle.co.jp/ultra_owa/ULTRA/usearch?rdo_1=1 �}�j���A��/���i��� ����
'  http://otn.oracle.co.jp/otn_pl/otn_tool/sqlconstraction SQL�\������
'  http://otn.oracle.co.jp/otn_pl/otn_tool/err_code_proc   �G���[�E���b�Z�[�W����
'  http://otn.oracle.co.jp/otn_pl/otn_tool2/search_forum   OTN�f������
'
'���̂����F
'  ���ږ��������{�ꉻ
'  �A�h�C����
'
