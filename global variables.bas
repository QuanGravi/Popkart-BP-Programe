Attribute VB_Name = "global_variables"
'��ģ�����ڶ���ȫ�ֱ���
Public blueconfirm() As Integer
Public redconfirm() As Integer
Public gamestart As Integer
Public labelclick As Integer
Public maprace(9, 2) As Integer
Public racecoun As Integer
Public RRRA(9) As Integer
Public SV As Integer
Public theme As Integer
Public recorder As Integer
Public MPI(9) As Integer 'map pick index��PICKͼ���
'��һ�ε���ʷ��¼
Public IE(9, 1 To 20) As Boolean
Public LV(1, 1 To 20) As Boolean
Public L10E(1 To 20) As Boolean
Public L12E(1 To 20) As Boolean
Public L10C(1 To 20) As Integer
Public L12C(1 To 20) As Integer
Public L6C(1 To 20) As String
Public rc(1 To 20) As Integer
Public MR(9, 2, 1 To 20) As Integer
Public LC(1 To 20) As Integer
Public PCoun As Integer

'�������
Public time As Double
Public Period As Double
Public StopTime As Double
Public TimeInterval As Double
Public Amplitude() As Double
Public InitialPosition() As Double

Public imagenumber As Integer 'һ�������п�����ʾ��image����
Public coun As Integer
'��ͼ����������, R������, I�������
Public ThemeName() As String
Public MapN() As Integer
Public MapN_R() As Integer: Public MapN_I() As Integer:
Public Maporder() As Integer
Public Maporder_R() As Integer: Public Maporder_I() As Integer:
Public Mapname() As String
Public Mapname_R() As String: Public Mapname_I() As String
Public all_map As Integer
Public all_map_R As Integer: Public all_map_I As Integer: 'ͼ����ȫ����ͼ������
Public ALLMAP As Integer
Public ALLMAP_R As Integer: Public ALLMAP_I As Integer 'ȫ������ͼ��ȫ������ͼ������
Public AllMapN() As Integer
Public AllMapN_R() As Integer: Public AllMapN_I() As Integer
Public AllMaporder() As Integer
Public AllMaporder_R() As Integer: Public AllMaporder_I() As Integer
Public PicturePath As String '����·��or����·��
'ӳ������
Public index2theme() As Integer 'ȫͼ��ʾ��image��ŵ�������ŵ�ӳ������
Public index2map() As Integer 'ȫͼ��ʾ��image��ŵ���ͼ��ŵ�ӳ������
Public thememap2index() As Integer '����ӳ������
Public allindex2mapindex() As Integer 'ȫͼ��ŵ�����ͼ��ŵ�ӳ������
Public thememap2index_1() As Integer '������ʾ�µ�ͼ��ŵ�image��ŵ�ӳ������
Public PositionAmount As Integer 'ȫͼ��ʾ�����ɵĵ�ͼ����
'bpͼ��ż�¼����
Public mapindex As Integer
Public mapbp(19, 2) As Integer
Public indexbp(120) As Boolean
Public thememapbp() As Boolean
'�����������
Public searchsum As Integer
Public searchmode As Integer
Public searchindex2theme(4) As Integer '����ͼ��ŵ�������ŵ�ӳ������
Public searchindex2map(4) As Integer '����ͼ��ŵ���ͼ��ŵ�ӳ������
Public thememap2searchindex() As Integer '����ӳ������
'����ʱ���
Public anticlock As Integer
Public waitingtime As Integer
'ͼ���������
Public map_pool_choice As Integer: Public map_pool_choice_I As Integer
'ͼ�ش������, ��׺I���������ͼ�ش������
Public thememapsave() As Boolean: Public thememapsaveI() As Boolean
Public theme_poolcreation As Integer: Public theme_poolcreationI As Integer
Public sum_for_scroll As Integer: Public sum_for_scrollI As Integer
Public pool_load_choice As Integer: Public pool_load_choiceI As Integer
Public PoolNumber As Integer: Public PoolNumberI As Integer
Public PoolRorNot As Boolean '��������ͼ��ʱ��ֵΪTure, ����Ϊfalse
Public maporder_save() As Integer
Public mapnumber_save() As Integer
Public th_map2index_pc() As Integer
Public current_pool_number As Integer
Public temname(17) As String
'BP�Զ���
Public BlueBPdata() As String
Public RedBPdata() As String
Public Type points_attribution '�ڵ����Խṹ��
    Turns As Integer
    TeamBPTypes As String
    StopOrNot As Boolean
    BeforeStop() As Integer
    AfterStop() As Integer
End Type
Public Type points_judge
    PreTeam As String
    NowTeam As String
    BlueStopOrNot As Boolean
    RedStopOrNot As Boolean
    BlueConfirmOrNot As Boolean
    RedConfirmOrNot As Boolean
End Type
Public PointAttriArray() As points_attribution '�ڵ���������
Public AllBPTurns As Integer 'BP�ܽڵ���
Public BP_choice() As String
Public FirstRoundNumber As Integer '������һ��BP�����Ľڵ���
Public BPnumber As Integer
Public BP_load_choice As Integer 'BP��������ѡ���BP�±�
'UI���
Public m_oCtlCancelMode As Object
