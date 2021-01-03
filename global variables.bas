Attribute VB_Name = "global_variables"
'此模块用于定义全局变量
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
Public MPI(9) As Integer 'map pick index，PICK图序号
'上一次的历史记录
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

'动画相关
Public time As Double
Public Period As Double
Public StopTime As Double
Public TimeInterval As Double
Public Amplitude() As Double
Public InitialPosition() As Double

Public imagenumber As Integer '一个窗体中可以显示的image数量
Public coun As Integer
'地图池设置数组, R代表竞速, I代表道具
Public ThemeName() As String
Public MapN() As Integer
Public MapN_R() As Integer: Public MapN_I() As Integer:
Public Maporder() As Integer
Public Maporder_R() As Integer: Public Maporder_I() As Integer:
Public Mapname() As String
Public Mapname_R() As String: Public Mapname_I() As String
Public all_map As Integer
Public all_map_R As Integer: Public all_map_I As Integer: '图池中全部地图的数量
Public ALLMAP As Integer
Public ALLMAP_R As Integer: Public ALLMAP_I As Integer '全部竞速图与全部道具图的数量
Public AllMapN() As Integer
Public AllMapN_R() As Integer: Public AllMapN_I() As Integer
Public AllMaporder() As Integer
Public AllMaporder_R() As Integer: Public AllMaporder_I() As Integer
Public PicturePath As String '道具路径or竞速路径
'映射数组
Public index2theme() As Integer '全图显示下image序号到主题序号的映射数组
Public index2map() As Integer '全图显示下image序号到地图序号的映射数组
Public thememap2index() As Integer '反向映射数组
Public allindex2mapindex() As Integer '全图序号到主题图序号的映射数组
Public thememap2index_1() As Integer '主题显示下地图序号到image序号的映射数组
Public PositionAmount As Integer '全图显示可容纳的地图数量
'bp图序号记录数组
Public mapindex As Integer
Public mapbp(19, 2) As Integer
Public indexbp(120) As Boolean
Public thememapbp() As Boolean
'搜索功能相关
Public searchsum As Integer
Public searchmode As Integer
Public searchindex2theme(4) As Integer '搜索图编号到主题序号的映射数组
Public searchindex2map(4) As Integer '搜索图编号到地图序号的映射数组
Public thememap2searchindex() As Integer '反向映射数组
'倒计时相关
Public anticlock As Integer
Public waitingtime As Integer
'图池设置相关
Public map_pool_choice As Integer: Public map_pool_choice_I As Integer
'图池创建相关, 后缀I代表与道具图池创建相关
Public thememapsave() As Boolean: Public thememapsaveI() As Boolean
Public theme_poolcreation As Integer: Public theme_poolcreationI As Integer
Public sum_for_scroll As Integer: Public sum_for_scrollI As Integer
Public pool_load_choice As Integer: Public pool_load_choiceI As Integer
Public PoolNumber As Integer: Public PoolNumberI As Integer
Public PoolRorNot As Boolean '操作竞速图池时该值为Ture, 否则为false
Public maporder_save() As Integer
Public mapnumber_save() As Integer
Public th_map2index_pc() As Integer
Public current_pool_number As Integer
Public temname(17) As String
'BP自定义
Public BlueBPdata() As String
Public RedBPdata() As String
Public Type points_attribution '节点属性结构体
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
Public PointAttriArray() As points_attribution '节点属性数组
Public AllBPTurns As Integer 'BP总节点数
Public BP_choice() As String
Public FirstRoundNumber As Integer '蓝方第一次BP包含的节点数
Public BPnumber As Integer
Public BP_load_choice As Integer 'BP创建界面选择的BP下标
'UI相关
Public m_oCtlCancelMode As Object
